import math
import openpyxl

wb = openpyxl.load_workbook("12627223.xlsx", data_only=True)
ws = wb["Daily Log"]

sleep = []
fitness = []
study = []
coding = []
class_time = []
class_attended = []
other_activities = []
total_tracked = []
free_uncounted = []
feeling_scores = []
sat_scores = []
energy_scores = []

f_map = {"excellent": 5, "good": 4, "neutral": 3, "stressed": 2, "low": 1}
s_map = {"very satisfied": 5, "satisfied": 4, "neutral": 3, "unsatisfied": 2, "very unsatisfied": 1}
e_map = {"high": 3, "medium": 2, "low": 1}

expected_days = 36
invalid_records = 0

for r in range(6, ws.max_row + 1):
    date_val = ws.cell(r, 1).value
    if date_val is None:
        continue

    try:
        s = float(ws.cell(r, 2).value or 0)
        f = float(ws.cell(r, 3).value or 0)
        st = float(ws.cell(r, 4).value or 0)
        c = float(ws.cell(r, 5).value or 0)
        cl = float(ws.cell(r, 6).value or 0)
        ca = float(ws.cell(r, 7).value or 0)
        ot = float(ws.cell(r, 8).value or 0)

        # Direct computation fallback if cached formula value is None
        tr_val = ws.cell(r, 9).value
        tr = float(tr_val) if tr_val is not None else (s + f + st + c + cl + ot)

        fr_val = ws.cell(r, 10).value
        fr = float(fr_val) if fr_val is not None else (1440.0 - tr)
    except (ValueError, TypeError):
        invalid_records += 1
        continue

    sleep.append(s)
    fitness.append(f)
    study.append(st)
    coding.append(c)
    class_time.append(cl)
    class_attended.append(ca)
    other_activities.append(ot)
    total_tracked.append(tr)
    free_uncounted.append(fr)

    f_txt = str(ws.cell(r, 11).value or "").strip().lower()
    s_txt = str(ws.cell(r, 12).value or "").strip().lower()
    e_txt = str(ws.cell(r, 13).value or "").strip().lower()

    feeling_scores.append(f_map.get(f_txt, 3))
    sat_scores.append(s_map.get(s_txt, 3))
    energy_scores.append(e_map.get(e_txt, 2))

valid_days = len(sleep)
missing_days = max(0, expected_days - valid_days)


def calculate_mean(lst):
    return sum(lst) / len(lst) if lst else 0.0

def calculate_correlation(x, y):
    n = len(x)
    if n == 0 or len(y) != n:
        return 0.0

    mean_x = calculate_mean(x)
    mean_y = calculate_mean(y)

    numerator = sum((x[i] - mean_x) * (y[i] - mean_y) for i in range(n))
    sum_sq_x = sum((x[i] - mean_x) ** 2 for i in range(n))
    sum_sq_y = sum((y[i] - mean_y) ** 2 for i in range(n))

    denominator = math.sqrt(sum_sq_x * sum_sq_y)
    return numerator / denominator if denominator != 0 else 0.0


def calculate_tpi(coding):
    return calculate_mean(coding)

def calculate_aai(study, class_time):
    return calculate_mean([s + c for s, c in zip(study, class_time)])

def calculate_phai(fitness):
    return calculate_mean(fitness)

def calculate_sri(sleep):
    return calculate_mean(sleep)

def calculate_abi(free_uncounted):
    return calculate_mean(free_uncounted)

def calculate_tui(total_tracked):
    return calculate_mean(total_tracked)

def calculate_ei(feeling, sat, energy):
    combined = [f + s + e for f, s, e in zip(feeling, sat, energy)]
    return sum(combined) / (3 * len(combined)) if combined else 0.0

def calculate_dci(valid_days, expected_days=36):
    return min(100.0, (valid_days / expected_days) * 100) if expected_days else 0.0

def calculate_pai(tpi, aai, phai, sri, tui, ei, dci):
    return ( 0.15 * tpi + 0.20 * aai + 0.15 * phai + 0.20 * sri + 0.15 * tui + 0.10 * ei + 0.05 * dci )

avg_sleep = calculate_mean(sleep)
avg_fitness = calculate_mean(fitness)
avg_study = calculate_mean(study)
avg_coding = calculate_mean(coding)
avg_class = calculate_mean(class_time)
avg_other = calculate_mean(other_activities)
avg_free = calculate_mean(free_uncounted)

dci = calculate_dci(valid_days, expected_days)
tpi = calculate_tpi(coding)
aai = calculate_aai(study, class_time)
phai = calculate_phai(fitness)
sri = calculate_sri(sleep)
abi = calculate_abi(free_uncounted)
tui = calculate_tui(total_tracked)
ei = calculate_ei(feeling_scores, sat_scores, energy_scores)
pai = calculate_pai(tpi, aai, phai, sri, tui, ei, dci)

corr_coding_energy = calculate_correlation(coding, energy_scores)
corr_sleep_energy = calculate_correlation(sleep, energy_scores)
corr_study_sat = calculate_correlation(study, sat_scores)

print("1. Activity Data Summary")
print("-" * 55)
print(f"{'Particular':<36} | {'Value':<15}")
print("-" * 55)
print(f"{'Expected number of days':<36} | {expected_days} days")
print(f"{'Valid days recorded':<36} | {valid_days} days")
print(f"{'Missing days':<36} | {missing_days} days")
print(f"{'Invalid / excluded records':<36} | {invalid_records} records")
print(f"{'Average Sleep/day':<36} | {avg_sleep:.2f} min/day")
print(f"{'Average Fitness/day':<36} | {avg_fitness:.2f} min/day")
print(f"{'Average Study/day':<36} | {avg_study:.2f} min/day")
print(f"{'Average Coding/day':<36} | {avg_coding:.2f} min/day")
print(f"{'Average Class/day':<36} | {avg_class:.2f} min/day")
print(f"{'Average Other Activities/day':<36} | {avg_other:.2f} min/day")
print(f"{'Average Free / Unaccounted Time/day':<36} | {avg_free:.2f} min/day")
print("-" * 55)

print("\n" + "=" * 50)
print(f"{'ACTIVITY INDICES & RESULTS':^55}")
print("=" * 50)
print(f"Data Continuity (DCI)   : {dci:.2f}%")
print(f"Tech Productivity (TPI) : {tpi:.2f} min/day")
print(f"Academic Activity (AAI) : {aai:.2f} min/day")
print(f"Physical Activity (PhAI): {phai:.2f} min/day")
print(f"Sleep & Recovery (SRI)  : {sri:.2f} min/day")
print(f"Time Utilization (TUI)  : {tui:.2f} min/day")
print(f"Activity Balance (ABI)  : {abi:.2f} min/day")
print(f"Experience Index (EI)   : {ei:.2f} / 5.0")
print(f"Personal Activity Index : {pai:.2f}")
print("-" * 50)
print("Relationship Analysis (Correlations):")
print(f"  Coding vs Energy      : {corr_coding_energy:.4f}")
print(f"  Sleep vs Energy       : {corr_sleep_energy:.4f}")
print(f"  Study vs Satisfaction : {corr_study_sat:.4f}")
print("=" * 50)