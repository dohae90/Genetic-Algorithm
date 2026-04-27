

import os
import random
from flask import Flask

app = Flask(__name__)
DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL READER
# ─────────────────────────────────────────────────────────────────────────────

def load_excel(filename):
    """
    Read an xlsx file and return a list of row-dicts.
    Returns [] if file missing or openpyxl unavailable.
    """
    try:
        import openpyxl
        path = os.path.join(DATA_DIR, filename)
        if not os.path.exists(path):
            return []
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        if not rows:
            return []
        headers = [str(h).strip() if h is not None else '' for h in rows[0]]
        return [dict(zip(headers, r)) for r in rows[1:] if any(v is not None for v in r)]
    except Exception as e:
        print(f"[WARN] Cannot load {filename}: {e}")
        return []


# ─────────────────────────────────────────────────────────────────────────────
# BUILD R MATRIX  (user-item ratings)
# ─────────────────────────────────────────────────────────────────────────────

def build_R(ratings_raw, behavior_raw):
    """
    R[(user_id, product_id)] = rating in [1, 5].

    Sources (with priority):
      1. Explicit ratings from ratings.xlsx
      2. Implicit scores from behavior.xlsx:
           score = 1 + viewed*0.1 + clicked*0.5 + purchased*2.0  (capped at 5)
    """
    R = {}

    for row in ratings_raw:
        try:
            uid  = int(float(row['user_id']))
            pid  = int(float(row['product_id']))
            rate = float(row['rating'])
            if 1 <= rate <= 5:
                R[(uid, pid)] = rate
        except (ValueError, TypeError, KeyError):
            continue

    for row in behavior_raw:
        try:
            uid = int(float(row['user_id']))
            pid = int(float(row['product_id']))
            if (uid, pid) in R:
                continue  # explicit rating has priority
            v = float(row.get('viewed',    0) or 0)
            c = float(row.get('clicked',   0) or 0)
            p = float(row.get('purchased', 0) or 0)
            if v == 0 and c == 0 and p == 0:
                continue
            score = min(5.0, 1.0 + v * 0.1 + c * 0.5 + p * 2.0)
            R[(uid, pid)] = round(score, 2)
        except (ValueError, TypeError, KeyError):
            continue

    return R


# ─────────────────────────────────────────────────────────────────────────────
# BUILD P MATRIX  (item-tag, analogous to the paper's tag genome)
# ─────────────────────────────────────────────────────────────────────────────

def build_P(products, behavior_raw):
    """
    Construct the item-tag matrix P where every product gets a unique
    multi-dimensional vector in [0, 1].

    Tag dimensions (= columns of P):
      [0 .. N_cat-1]  : category membership  (0 or 1)  — one-hot
      [N_cat]         : normalised price      [0, 1]
      [N_cat+1]       : normalised view count [0, 1]   ← from behavior.xlsx
      [N_cat+2]       : normalised click rate [0, 1]
      [N_cat+3]       : normalised purchase rate [0, 1]

    This is the direct analogue of the tag-genome matrix in the paper:
      "rel(t, i) ∈ P — the relevance of item i to tag t, between 0 and 1"

    Returns:
      P         : dict { product_id : [float, ...] }
      tag_names : list of tag label strings (for display)
      categories: sorted list of category strings
    """
    # Aggregate product-level behaviour totals
    views     = {}
    clicks    = {}
    purchases = {}
    for row in behavior_raw:
        try:
            pid = int(float(row['product_id']))
            views[pid]     = views.get(pid, 0)     + float(row.get('viewed',    0) or 0)
            clicks[pid]    = clicks.get(pid, 0)    + float(row.get('clicked',   0) or 0)
            purchases[pid] = purchases.get(pid, 0) + float(row.get('purchased', 0) or 0)
        except (ValueError, TypeError, KeyError):
            continue

    # Sorted category list determines the one-hot positions
    categories = sorted({str(p.get('category', 'Unknown')).strip()
                         for p in products.values()})
    cat_index  = {c: i for i, c in enumerate(categories)}

    # Normalisation denominators
    all_prices    = [float(p.get('price', 0) or 0) for p in products.values()]
    max_price     = max(all_prices) or 1.0
    max_views     = max(views.values(),     default=1) or 1.0
    max_clicks    = max(clicks.values(),    default=1) or 1.0
    max_purchases = max(purchases.values(), default=1) or 1.0

    P = {}
    for pid, prod in products.items():
        cat = str(prod.get('category', 'Unknown')).strip()

        # ── Tag features ──────────────────────────────────────────────
        cat_vec  = [1.0 if c == cat else 0.0 for c in categories]  # one-hot category
        price_f  = float(prod.get('price', 0) or 0) / max_price    # price level
        view_f   = views.get(pid, 0)     / max_views                # popularity: views
        click_f  = clicks.get(pid, 0)    / max_clicks               # popularity: clicks
        purch_f  = purchases.get(pid, 0) / max_purchases            # popularity: purchases

        P[pid] = cat_vec + [price_f, view_f, click_f, purch_f]

    tag_names = categories + ['Price-Level', 'View-Count', 'Click-Rate', 'Purchase-Rate']
    return P, tag_names, categories


# ─────────────────────────────────────────────────────────────────────────────
# MAIN DATA LOADER
# ─────────────────────────────────────────────────────────────────────────────

def load_data():
    """
    Load all four xlsx files and build R and P matrices.
    Falls back to built-in sample data if files are missing.

    Returns: (users, products, R, P, tag_names, categories, using_sample)
    """
    users_raw    = load_excel('users.xlsx')
    products_raw = load_excel('products.xlsx')
    ratings_raw  = load_excel('ratings.xlsx')
    behavior_raw = load_excel('behavior.xlsx')

    if not (users_raw and products_raw and ratings_raw):
        return build_sample_data()

    # Parse users
    users = {}
    for row in users_raw:
        try:
            uid = int(float(row['user_id']))
            users[uid] = {'user_id': uid,
                          'age':      row.get('age', '?'),
                          'location': row.get('location', '?')}
        except (ValueError, TypeError, KeyError):
            continue

    # Parse products
    products = {}
    for row in products_raw:
        try:
            pid = int(float(row['product_id']))
            products[pid] = {'product_id': pid,
                             'category':   str(row.get('category', 'Unknown')).strip(),
                             'price':      float(row.get('price', 0) or 0)}
        except (ValueError, TypeError, KeyError):
            continue

    R                          = build_R(ratings_raw, behavior_raw)
    P, tag_names, categories   = build_P(products, behavior_raw)

    return users, products, R, P, tag_names, categories, False


def build_sample_data():
    """
    Built-in demo data (10 users, 20 products, 5 categories).
    Each product has a unique P vector via varied prices and behaviour.
    """
    random.seed(99)

    categories = ['Electronics', 'Clothing', 'Books', 'Home', 'Sports']
    users      = {i: {'user_id': i, 'age': 18 + i * 2, 'location': f'City-{i}'}
                  for i in range(1, 11)}

    products = {}
    for pid in range(1, 21):
        cat = categories[(pid - 1) % len(categories)]
        # Different prices within the same category → unique P vectors
        products[pid] = {'product_id': pid,
                         'category':   cat,
                         'price':      round(5.0 + pid * 4.7 + (pid % 3) * 8, 2)}

    # Simulate behavior (each user interacted with ~10 products)
    random.seed(77)
    behavior_raw = []
    for uid in range(1, 11):
        for pid in random.sample(range(1, 21), 10):
            purchased = 1 if random.random() < 0.3 else 0
            clicked   = 1 if (purchased or random.random() < 0.5) else 0
            viewed    = random.randint(1, 5)
            behavior_raw.append({'product_id': pid, 'viewed': viewed,
                                  'clicked': clicked, 'purchased': purchased})

    P, tag_names, cats = build_P(products, behavior_raw)

    # Build R from explicit (simulated) ratings
    random.seed(42)
    R = {}
    for uid in range(1, 11):
        for pid in random.sample(range(1, 21), 7):
            R[(uid, pid)] = round(random.uniform(1.5, 5.0), 1)

    return users, products, R, P, tag_names, cats, True



if __name__ == '__main__':
    app.run(debug=True, port=5000)
