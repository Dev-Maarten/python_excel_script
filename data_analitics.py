import pandas as pd

df = pd.read_excel("Overzicht stemmen.xlsx")
df.columns = df.columns.str.strip()


stem_count_col = df.columns[2]  

stemkolommen = df.columns[5:12]


for col in stemkolommen:
    df[col] = df[col].astype(str).str.strip().str.upper()

onderwerpen = {
    0: "Laadinfrastructuur",
    1: "Compensatie",
    2: "Eenmalig",
    3: "Lening",
    4: "Offerte",
    5: "Uitzetten",
    6: "Lening aanvraag"
}

stemverdeling = []

for i, col in enumerate(stemkolommen):
    onderwerp = onderwerpen.get(i, f"Onderwerp {i}")

    stemmen_v = df.loc[df[col] == "V", stem_count_col].sum()
    stemmen_t = df.loc[df[col] == "T", stem_count_col].sum()
    stemmen_m = df.loc[df[col] == "M", stem_count_col].sum()

    if stemmen_v >= stemmen_t:
        stemmen_v += stemmen_m
    else:
        stemmen_t += stemmen_m

    totaal = stemmen_v + stemmen_t

    pct_v = (stemmen_v / totaal * 100) if totaal > 0 else 0
    pct_t = (stemmen_t / totaal * 100) if totaal > 0 else 0

    stemverdeling.append({
        "Onderwerp": onderwerp,
        "Stemmen Voor (incl. I)": stemmen_v,
        "Stemmen Tegen (incl. I)": stemmen_t,
        "Totaal": totaal,
        "Percentage Voor": round(pct_v, 2),
        "Percentage Tegen": round(pct_t, 2)
    })

result_df = pd.DataFrame(stemverdeling)
result_df.to_excel("Stemresultaten_analyse.xlsx", sheet_name="Analyse Per Onderwerp", index=False)