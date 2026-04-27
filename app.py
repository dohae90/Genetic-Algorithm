import os
import math
import random
from flask import Flask, render_template, request, jsonify

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


# ─────────────────────────────────────────────────────────────────────────────
# MATRIX HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def dot(a, b):
    """Dot product of two equal-length lists."""
    return sum(x * y for x, y in zip(a, b))


def compute_rmse(Q, P, R, user_ids):
    """
    RMSE between actual ratings R and predicted ratings R̂ = Q × Pᵀ.

    Q[i] is user i's tag-affinity row.
    P[pid] is product pid's tag vector (values in [0, 1]).
    Only (uid, pid) pairs present in R are evaluated.

    Scale alignment (key fix):
      R is in [1, 5]; Q·Pᵀ is in [0, 1] (since Q and P both ∈ [0,1]).
      With only 11 tag dimensions, Q·Pᵀ cannot reach [1,5].
      We normalise R to [0,1]: R_norm = (R - 1) / 4
      This brings both sides to the same scale so RMSE is meaningful.
      The paper avoids this issue by using 1128 tags whose cumulative
      dot product naturally reaches [1, 5] range.
    """
    uid_idx = {uid: i for i, uid in enumerate(user_ids)}
    sse, n = 0.0, 0

    for (uid, pid), actual in R.items():
        if uid not in uid_idx or pid not in P:
            continue
        actual_norm = (actual - 1.0) / 4.0
        predicted = dot(Q[uid_idx[uid]], P[pid])
        sse += (actual_norm - predicted) ** 2
        n += 1

    return math.sqrt(sse / n) if n > 0 else float('inf')


# ─────────────────────────────────────────────────────────────────────────────
# GENETIC ALGORITHM — ETagMF  (Algorithm 1 from the paper)
# ─────────────────────────────────────────────────────────────────────────────

def init_population(Np, n_users, n_tags):
    """
    Generate Np random chromosomes.
    Each chromosome = Q matrix (n_users × n_tags), genes ∈ [0, 1].
    (Equation initialisation from Section III-A of the paper.)
    """
    return [[[random.random() for _ in range(n_tags)]
             for _ in range(n_users)]
            for _ in range(Np)]


def roulette_select(population, fitnesses):
    """
    Roulette-Wheel Parent Selection (RWPSA — Section III-A).
    Lower RMSE → higher selection probability  (weight = 1/fitness).
    """
    weights = [1.0 / (f + 1e-10) for f in fitnesses]
    total = sum(weights)
    pick = random.random() * total
    cumul = 0.0
    for chrom, w in zip(population, weights):
        cumul += w
        if cumul >= pick:
            return chrom
    return population[-1]


def crossover(par1, par2, Pc):
    """
    Single-row crossover (Section III-B of the paper, Example 2).
    With probability Pc, one randomly chosen row is swapped between parents.
    """
    if random.random() > Pc:
        return [r[:] for r in par1], [r[:] for r in par2]

    idx = random.randint(0, len(par1) - 1)
    child1 = [par2[i][:] if i == idx else par1[i][:] for i in range(len(par1))]
    child2 = [par1[i][:] if i == idx else par2[i][:] for i in range(len(par2))]
    return child1, child2


def mutate(chrom, Pm):
    """
    Gene-wise mutation (Section III-C, Example 3).
    Each gene is replaced by a new random value in [0, 1] with probability Pm.
    """
    return [[random.random() if random.random() < Pm else g
             for g in row]
            for row in chrom]


def run_ga(P, R, user_ids, Np=20, Ni=100, Pc=0.8, Pm=0.05, seed=42):
    """
    ETagMF main loop (Algorithm 1 from the paper).

    What this does (step by step, matching the paper):
      1. Initialize population of Np chromosomes (random Q matrices)
      2. Evaluate fitness of each chromosome via RMSE(R, Q×Pᵀ)
      3. For each generation:
           a. Keep best chromosome (elitism)
           b. Select two parents by Roulette Wheel
           c. Apply row-crossover with probability Pc
           d. Apply gene-mutation with probability Pm
           e. Evaluate children and insert better ones
      4. Return best Q found + RMSE history per generation

    Parameters match the paper's notation:
      Np = population size
      Ni = number of generations (MaxGeneration)
      Pc = crossover probability  (paper: 0.8)
      Pm = mutation probability   (paper: 0.5 — we use smaller for stability)
    """
    random.seed(seed)
    n_users = len(user_ids)
    n_tags = len(next(iter(P.values()))) if P else 1

    population = init_population(Np, n_users, n_tags)
    fitnesses = [compute_rmse(ch, P, R, user_ids) for ch in population]

    best_i = min(range(Np), key=lambda i: fitnesses[i])
    best_Q = [row[:] for row in population[best_i]]
    best_rmse = fitnesses[best_i]
    history = []

    for _ in range(Ni):
        next_pop = [[row[:] for row in best_Q]]
        next_fit = [best_rmse]

        while len(next_pop) < Np:
            p1 = roulette_select(population, fitnesses)
            p2 = roulette_select(population, fitnesses)

            c1, c2 = crossover(p1, p2, Pc)
            c1 = mutate(c1, Pm)
            c2 = mutate(c2, Pm)

            for child in [c1, c2]:
                if len(next_pop) < Np:
                    f = compute_rmse(child, P, R, user_ids)
                    next_pop.append(child)
                    next_fit.append(f)

        population = next_pop
        fitnesses = next_fit

        gen_best_i = min(range(Np), key=lambda i: fitnesses[i])
        if fitnesses[gen_best_i] < best_rmse:
            best_rmse = fitnesses[gen_best_i]
            best_Q = [row[:] for row in population[gen_best_i]]

        history.append(round(best_rmse, 5))

    return best_Q, history


# ─────────────────────────────────────────────────────────────────────────────
# RECOMMENDATION  (R̂ = Q × Pᵀ, top-N unrated items)
# ─────────────────────────────────────────────────────────────────────────────

def recommend(Q, P, R, user_ids, products, uid, top_n=5):
    """
    For the target user (uid), predict ratings for all unrated products
    using R̂[uid][pid] = Q[uid_row] · P[pid], then return top-N.
    """
    uid_idx = {u: i for i, u in enumerate(user_ids)}
    if uid not in uid_idx:
        return []

    q_row = Q[uid_idx[uid]]
    already = {pid for (u, pid) in R if u == uid}

    scores = []
    for pid, p_vec in P.items():
        if pid in already:
            continue
        scores.append((pid, round(dot(q_row, p_vec), 4)))

    scores.sort(key=lambda x: -x[1])
    result = []
    for pid, score in scores[:top_n]:
        info = products.get(pid, {})
        result.append({'product_id': pid,
                       'category': str(info.get('category', '?')),
                       'price': float(info.get('price', 0)),
                       'predicted': score})
    return result


# ─────────────────────────────────────────────────────────────────────────────
# FLASK ROUTES
# ─────────────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    users, products, R, P, tag_names, categories, using_sample = load_data()
    return render_template(
        'index.html',
        user_ids=sorted(users.keys()),
        tag_names=tag_names,
        categories=categories,
        num_users=len(users),
        num_products=len(products),
        num_ratings=len(R),
        num_tags=len(tag_names),
        using_sample=using_sample,
    )


@app.route('/explain')
def explain():
    """Detailed algorithm explanation page."""
    return render_template('explain.html')


@app.route('/paper-to-code')
def paper_to_code():
    """Paper → Code translation walkthrough page."""
    return render_template('paper_to_code.html')


@app.route('/run', methods=['POST'])
def run_endpoint():
    """
    POST JSON:
      user_id, top_n, Np, Ni, Pc, Pm, seed

    Returns JSON:
      rmse_history, final_rmse, recommendations,
      user_ratings, q_values, tag_names, user_info
    """
    body = request.get_json(force=True)
    uid = int(body.get('user_id', 1))
    top_n = max(1, int(body.get('top_n', 5)))
    Np = max(5, int(body.get('Np', 20)))
    Ni = max(10, int(body.get('Ni', 80)))
    Pc = float(body.get('Pc', 0.8))
    Pm = float(body.get('Pm', 0.05))
    seed = int(body.get('seed', 42))

    users, products, R, P, tag_names, categories, using_sample = load_data()
    user_ids = sorted(users.keys())

    best_Q, history = run_ga(P, R, user_ids,
                             Np=Np, Ni=Ni, Pc=Pc, Pm=Pm, seed=seed)

    recs = recommend(best_Q, P, R, user_ids, products, uid, top_n)

    user_ratings = []
    for (u, pid), rating in R.items():
        if u != uid:
            continue
        info = products.get(pid, {})
        user_ratings.append({'product_id': pid,
                             'category': str(info.get('category', '?')),
                             'price': float(info.get('price', 0)),
                             'rating': rating})
    user_ratings.sort(key=lambda x: -x['rating'])

    uid_idx = {u: i for i, u in enumerate(user_ids)}
    q_values = []
    if uid in uid_idx:
        q_row = best_Q[uid_idx[uid]]
        q_values = [{'tag': t, 'affinity': round(v, 4)}
                    for t, v in zip(tag_names, q_row)]

    return jsonify({
        'rmse_history': history,
        'final_rmse': history[-1] if history else 0,
        'recommendations': recs,
        'user_ratings': user_ratings[:10],
        'q_values': q_values,
        'tag_names': tag_names,
        'user_info': {k: str(v) for k, v in (users.get(uid) or {}).items()},
        'using_sample': using_sample,
    })


if __name__ == '__main__':
    app.run(debug=True, port=5000)
