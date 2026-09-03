from pathlib import Path
import json, re, sys
import pandas as pd

BASE=Path(__file__).resolve().parent
CURRENT=BASE/"Book_current.xlsx"
PREVIOUS=BASE/"Book_previous.xlsx"
TEMPLATE=BASE/"pardosana_vynoteka_2weeks_template.html"
OUTPUT=BASE/"pardosana_vynoteka_2weeks.html"
VAT=1.21

def norm_name(x):
    s=re.sub(r"\s+"," ",str(x).strip()).lower()
    return re.sub(r"\s+\d+(?:[,.]\d+)?%\s*$","",s)

BASE_COLUMNS={"Kods","EAN kods","Nosaukums","Produktų skyrius","Prekės kategorija","Prekės grupė"}

def clean_code(x):
    s=str(x).strip()
    return "" if s.lower()=="nan" else s

def clean_header(x):
    return re.sub(r"\s+"," ",str(x).strip())

def read_prices(st):
    pr_header=None
    for i in range(len(st)):
        if clean_header(st.iloc[i,0])=="Produkta nosaukums":
            pr_header=i
            break
    if pr_header is None:
        raise ValueError("Iestatījumi: nav cenu tabulas")

    price_rows=[]
    for i in range(pr_header+1,len(st)):
        n=clean_header(st.iloc[i,0])
        if not n or n.lower()=="nan":
            continue
        try:
            price_rows.append((n,float(st.iloc[i,1]),float(st.iloc[i,2])))
        except (TypeError,ValueError):
            continue
    return {norm_name(n):(s,p) for n,s,p in price_rows}

def read_current(path, store_order=None):
    df=pd.read_excel(path,sheet_name="Sheet1")
    st=pd.read_excel(path,sheet_name="Iestatījumi",header=None)
    df.columns=[clean_header(c) for c in df.columns]

    required={"Kods","EAN kods","Nosaukums","Kopā"}
    missing=required-set(df.columns)
    if missing:
        raise ValueError(f"{path.name}: trūkst kolonnas: {', '.join(sorted(missing))}")

    price_map=read_prices(st)
    detected_store_cols=[c for c in df.columns if c not in BASE_COLUMNS]

    if store_order is not None:
        by_header={clean_header(c):c for c in detected_store_cols}
        missing_stores=[c for c in store_order if c not in by_header]
        if missing_stores:
            raise ValueError(
                f"{path.name}: nav atrasti Book_previous veikali: "+", ".join(missing_stores)
            )
        store_cols=[by_header[c] for c in store_order]
    else:
        store_cols=detected_store_cols

    products=[]
    for _,r in df.iterrows():
        code=clean_code(r["Kods"])
        if not code:
            continue
        name=clean_header(r["Nosaukums"])
        key=norm_name(name)
        if key not in price_map:
            raise ValueError(
                f"{path.name}: cena nav atrasta produktam '{name}'. "
                "Pārbaudi Iestatījumi → Produkta nosaukums."
            )
        shelf,purchase=price_map[key]
        sell=shelf/VAT

        stores=[]
        for c in store_cols:
            v=pd.to_numeric(r[c],errors="coerce")
            sold=0 if pd.isna(v) else int(v)
            stores.append({
                "address":clean_header(c),
                "short":re.sub(r"\s+"," ",clean_header(c)).split(",")[0],
                "sold":sold
            })

        total_v=pd.to_numeric(r["Kopā"],errors="coerce")
        total=0 if pd.isna(total_v) else int(total_v)

        products.append({
            "code":code,"ean":clean_code(r["EAN kods"]),"name":name,"total":total,
            "shelf":shelf,"purchase":purchase,"sell_net":sell,
            "unit_profit":sell-purchase,"profit":total*(sell-purchase),"stores":stores
        })

    period_r=st[st[0].astype(str).str.strip().eq("Periods")]
    period=str(period_r.iloc[0,1]).strip() if not period_r.empty else path.stem
    return {"period":period,"products":products,"store_cols":store_cols}

def read_previous_4weeks(path):
    df=pd.read_excel(path,sheet_name="Sheet1")
    st=pd.read_excel(path,sheet_name="Iestatījumi",header=None)
    df.columns=[clean_header(c) for c in df.columns]
    price_map=read_prices(st)

    # Book_previous: one column per store containing the total sales for the last 4 weeks.
    # Expected header format: "<store> | 4 nedēļas kopā"
    suffix=" | 4 nedēļas kopā"
    store_cols=[]
    for c in df.columns:
        c_clean=clean_header(c)
        # Accept the 4-week store columns and ignore the aggregate "Kopā"
        # and "4 nedēļu vidējais" columns.
        if c_clean.endswith(suffix):
            store_name=c_clean[:-len(suffix)].strip()
            if store_name.lower() != "kopā":
                store_cols.append(c_clean)

    if not store_cols:
        # Fallback: the workbook may have store columns without the suffix.
        # Use all columns after the 6 product-identification columns, excluding
        # aggregate/calculation columns. This keeps the script compatible with
        # the same workbook even if the store headers were renamed.
        excluded = {
            "Kopā", "4 nedēļu vidējais", "Periods",
            "Kopā | 4 nedēļas kopā"
        }
        candidates=[]
        for c in df.columns[6:]:
            c_clean=clean_header(c)
            if not c_clean or c_clean in excluded:
                continue
            if c_clean.lower().startswith("4 nedēļu"):
                continue
            candidates.append(c_clean)
        store_cols=candidates

    if not store_cols:
        available=[clean_header(c) for c in df.columns if clean_header(c)]
        raise ValueError(
            "Book_previous.xlsx: nav atrastas veikalu kolonnas. "
            "Atrastās kolonnas: " + ", ".join(available)
        )

    stores=[c[:-len(suffix)].strip() for c in store_cols]

    products=[]
    for _,r in df.iterrows():
        code=clean_code(r["Kods"])
        if not code:
            continue
        name=clean_header(r["Nosaukums"])
        key=norm_name(name)
        if key not in price_map:
            raise ValueError(
                f"{path.name}: cena nav atrasta produktam '{name}'. "
                "Pārbaudi Iestatījumi → Produkta nosaukums."
            )
        shelf,purchase=price_map[key]
        sell=shelf/VAT

        stores_out=[]
        total_four_weeks=0.0
        for col,store in zip(store_cols,stores):
            v=pd.to_numeric(r[col],errors="coerce")
            four_week_total=0.0 if pd.isna(v) else float(v)
            avg_week=four_week_total/4.0
            total_four_weeks += four_week_total
            stores_out.append({
                "address":store,
                "short":re.sub(r"\s+"," ",store).split(",")[0],
                "sold":avg_week,
                "four_week_total":four_week_total
            })

        avg_total=total_four_weeks/4.0
        products.append({
            "code":code,"ean":clean_code(r["EAN kods"]),"name":name,"total":avg_total,
            "shelf":shelf,"purchase":purchase,"sell_net":sell,
            "unit_profit":sell-purchase,"profit":avg_total*(sell-purchase),
            "stores":stores_out
        })

    return {
        "period":"pēdējo 4 nedēļu vidējais",
        "products":products,
        "store_cols":stores
    }

def main():
    prev=read_previous_4weeks(PREVIOUS)
    cur=read_current(CURRENT, store_order=prev["store_cols"])
    pm={p["code"]:p for p in prev["products"]}
    items=[{"code":c["code"],"current":c,"previous":pm[c["code"]]}
           for c in cur["products"] if c["code"] in pm]
    if not items:
        raise ValueError("Nav vienādu produktu pēc Kods.")

    data={
        "current_period":cur["period"],
        "previous_period":prev["period"],
        "products":items
    }

    html=TEMPLATE.read_text(encoding="utf-8")
    html=html.replace("__DATA__",json.dumps(data,ensure_ascii=False,separators=(",",":")))

    legend="P1 — #Cēsu Premium pint alus can 0,568L 5% · P2 — Chiara Tuncis saulespuķu eļļā 80gx3"
    html=html.replace(
        "Pārdošana veikalos · visi produkti",
        f'Pārdošana veikalos · visi produkti<br><span style="font-size:9px;color:var(--muted);font-weight:600">{legend}</span>',
        1
    )

    OUTPUT.write_text(html,encoding="utf-8")
    print("Gatavs:",OUTPUT)

if __name__=="__main__":
    main()
