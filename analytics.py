"""Portfolio and ticket-level analytics for Daily Sales Register."""

from __future__ import annotations

import sqlite3
from typing import Optional, Sequence

from ingest import _merge_where, _schema_from_meta, compose_where, open_db
from schema import LABEL_TO_SQL

YES_TOKENS = ("yes", "y", "true", "1", "diamond", "international")


def compute_report_analytics(
    db_path: str,
    search: str = "",
    invoice_numbers: Optional[Sequence[str]] = None,
    filters: Optional[dict] = None,
) -> dict:
    con = open_db(db_path)
    try:
        meta = {r[0]: r[1] for r in con.execute("SELECT key, value FROM meta")}
        labels, sql_cols, label_to_sql = _schema_from_meta(meta)
        if invoice_numbers:
            ids = _clean_ids(invoice_numbers)
            if not ids:
                return empty_analytics(scope="selected")
            inv_col = label_to_sql.get("Invoice Number") or LABEL_TO_SQL.get("Invoice Number")
            if not inv_col:
                return empty_analytics(scope="selected")
            placeholders = ",".join(["?"] * len(ids))
            where = f'WHERE "{inv_col}" IN ({placeholders})'
            params = ids
            scope = "selected"
        else:
            where, params = compose_where(search, sql_cols, label_to_sql, filters)
            scope = "filtered" if where else "all"
        payload = _aggregate(con, label_to_sql, where, params)
        payload["scope"] = scope
        payload["search"] = (search or "").strip()
        payload["selectedCount"] = len(_clean_ids(invoice_numbers)) if invoice_numbers else payload["invoices"]
        return payload
    finally:
        con.close()


def compute_from_rows(rows: Sequence[dict]) -> dict:
    """Same metric definitions, used when selected tickets have no invoice number."""
    rows = [r for r in (rows or []) if isinstance(r, dict)]
    if not rows:
        return empty_analytics(scope="selected")

    def num(row, key):
        val = row.get(key)
        if val is None or val == "":
            return 0.0
        try:
            return float(val)
        except (TypeError, ValueError):
            return 0.0

    def text(row, key):
        return str(row.get(key) or "").strip()

    invoices = len(rows)
    billed = sum(num(r, "Invoice amount") for r in rows)
    due = sum(num(r, "Total Amount Due") for r in rows)
    sgst = sum(num(r, "SGST") for r in rows)
    cgst = sum(num(r, "CGST") for r in rows)
    igst = sum(num(r, "IGST") for r in rows)
    tds = sum(num(r, "TDS") for r in rows)
    charges = sum(num(r, "Bank Charges") for r in rows)
    cheque = sum(num(r, "Cheque Amt") for r in rows)
    usd = sum(num(r, "USD $ Billing value") for r in rows)
    fx_w = 0.0
    fx_d = 0.0
    risky_n = risky_due = force_n = force_amt = 0
    intl = diamond = realized = 0
    advertisers, brands, teams, managers, currencies = set(), set(), set(), set(), set()

    for r in rows:
        rate = num(r, "USD Booking $ Rate")
        usd_v = num(r, "USD $ Billing value")
        if usd_v:
            fx_w += rate * usd_v
            fx_d += usd_v
        flag = text(r, "Force Billings and Risky Debtors").lower()
        inv = num(r, "Invoice amount")
        due_r = num(r, "Total Amount Due")
        if "risk" in flag:
            risky_n += 1
            risky_due += due_r
        if "force" in flag:
            force_n += 1
            force_amt += inv
        if _is_yes(text(r, "International client")):
            intl += 1
        if _is_yes(text(r, "Diamond Seller")):
            diamond += 1
        if text(r, "Bank Realization date"):
            realized += 1
        if text(r, "Advertiser Name"):
            advertisers.add(text(r, "Advertiser Name"))
        if text(r, "Brand Name"):
            brands.add(text(r, "Brand Name"))
        if text(r, "Team Responsible"):
            teams.add(text(r, "Team Responsible"))
        if text(r, "Accounts Manager"):
            managers.add(text(r, "Accounts Manager"))
        if text(r, "Currency"):
            currencies.add(text(r, "Currency"))

    gst = sgst + cgst + igst
    collected_est = max(billed - due, 0.0)
    return _finalize(
        {
            "invoices": invoices,
            "advertisers": len(advertisers),
            "brands": len(brands),
            "teams": len(teams),
            "managers": len(managers),
            "currencies": len(currencies),
            "invoiceAmount": billed,
            "avgInvoice": billed / invoices if invoices else 0,
            "amountDue": due,
            "gst": gst,
            "sgst": sgst,
            "cgst": cgst,
            "igst": igst,
            "tds": tds,
            "bankCharges": charges,
            "chequeAmount": cheque,
            "usdBilling": usd,
            "weightedFx": (fx_w / fx_d) if fx_d else 0,
            "riskyCount": risky_n,
            "riskyDue": risky_due,
            "forceCount": force_n,
            "forceAmount": force_amt,
            "internationalCount": intl,
            "diamondCount": diamond,
            "realizedCount": realized,
            "unrealizedCount": invoices - realized,
            "collectedEst": collected_est,
            "topAdvertisers": _top_from_rows(rows, "Advertiser Name"),
            "topTeams": _top_from_rows(rows, "Team Responsible"),
            "topManagers": _top_from_rows(rows, "Accounts Manager"),
            "currencyMix": _top_from_rows(rows, "Currency"),
            "topOutstanding": _top_from_rows(rows, "Advertiser Name", limit=8, amount_key="Total Amount Due"),
            "topBrands": _top_from_rows(rows, "Brand Name"),
            "openDueCount": sum(1 for r in rows if num(r, "Total Amount Due") > 0.009),
            "unrealizedDue": round(sum(num(r, "Total Amount Due") for r in rows if not text(r, "Bank Realization date")), 2),
            "openInvoices": _open_from_rows(rows),
        },
        scope="selected",
    )


def empty_analytics(scope: str = "all") -> dict:
    return _finalize(
        {
            "invoices": 0,
            "advertisers": 0,
            "brands": 0,
            "teams": 0,
            "managers": 0,
            "currencies": 0,
            "invoiceAmount": 0,
            "avgInvoice": 0,
            "amountDue": 0,
            "gst": 0,
            "sgst": 0,
            "cgst": 0,
            "igst": 0,
            "tds": 0,
            "bankCharges": 0,
            "chequeAmount": 0,
            "usdBilling": 0,
            "weightedFx": 0,
            "riskyCount": 0,
            "riskyDue": 0,
            "forceCount": 0,
            "forceAmount": 0,
            "internationalCount": 0,
            "diamondCount": 0,
            "realizedCount": 0,
            "unrealizedCount": 0,
            "collectedEst": 0,
            "topAdvertisers": [],
            "topTeams": [],
            "topManagers": [],
            "currencyMix": [],
            "topOutstanding": [],
            "topBrands": [],
            "openDueCount": 0,
            "unrealizedDue": 0,
            "openInvoices": [],
        },
        scope=scope,
    )


def _aggregate(con: sqlite3.Connection, label_to_sql: dict, where: str, params: list) -> dict:
    def ident(label: str) -> Optional[str]:
        sql = label_to_sql.get(label)
        if not sql:
            return None
        return '"' + str(sql).replace('"', "") + '"'

    def coalesce_sum(label: str) -> str:
        col = ident(label)
        return f"COALESCE(SUM({col}), 0)" if col else "0"

    def coalesce_avg(label: str) -> str:
        col = ident(label)
        return f"COALESCE(AVG({col}), 0)" if col else "0"

    def distinct_count(label: str) -> str:
        col = ident(label)
        if not col:
            return "0"
        return (
            f"COUNT(DISTINCT CASE WHEN {col} IS NULL OR TRIM(CAST({col} AS TEXT)) IN ('', 'None') "
            f"THEN NULL ELSE {col} END)"
        )

    def filled(label: str) -> str:
        col = ident(label)
        if not col:
            return "0"
        return (
            f"(SUM(CASE WHEN {col} IS NOT NULL AND TRIM(CAST({col} AS TEXT)) NOT IN ('', 'None') "
            f"THEN 1 ELSE 0 END))"
        )

    risk = ident("Force Billings and Risky Debtors")
    due = ident("Total Amount Due")
    billed = ident("Invoice amount")
    intl = ident("International client")
    diamond = ident("Diamond Seller")
    usd = ident("USD $ Billing value")
    fx = ident("USD Booking $ Rate")

    risk_count = (
        f"SUM(CASE WHEN LOWER(CAST({risk} AS TEXT)) LIKE '%risk%' THEN 1 ELSE 0 END)"
        if risk else "0"
    )
    risk_due = (
        f"SUM(CASE WHEN LOWER(CAST({risk} AS TEXT)) LIKE '%risk%' THEN COALESCE({due}, 0) ELSE 0 END)"
        if risk and due else "0"
    )
    force_count = (
        f"SUM(CASE WHEN LOWER(CAST({risk} AS TEXT)) LIKE '%force%' THEN 1 ELSE 0 END)"
        if risk else "0"
    )
    force_amt = (
        f"SUM(CASE WHEN LOWER(CAST({risk} AS TEXT)) LIKE '%force%' THEN COALESCE({billed}, 0) ELSE 0 END)"
        if risk and billed else "0"
    )
    yes_expr = "IN ('yes','y','true','1','diamond','international')"
    intl_count = (
        f"SUM(CASE WHEN LOWER(TRIM(CAST({intl} AS TEXT))) {yes_expr} THEN 1 ELSE 0 END)"
        if intl else "0"
    )
    diamond_count = (
        f"SUM(CASE WHEN LOWER(TRIM(CAST({diamond} AS TEXT))) {yes_expr} THEN 1 ELSE 0 END)"
        if diamond else "0"
    )
    fx_num = (
        f"COALESCE(SUM(COALESCE({fx}, 0) * COALESCE({usd}, 0)), 0)"
        if fx and usd else "0"
    )
    fx_den = coalesce_sum("USD $ Billing value")
    bank_date = ident("Bank Realization date")
    open_due_expr = (
        f"SUM(CASE WHEN COALESCE({due}, 0) > 0.009 THEN 1 ELSE 0 END)" if due else "0"
    )
    unrealized_due_expr = (
        f"(SUM(CASE WHEN ({bank_date} IS NULL OR TRIM(CAST({bank_date} AS TEXT)) IN ('','None')) "
        f"THEN COALESCE({due}, 0) ELSE 0 END))"
        if due and bank_date else (f"COALESCE(SUM({due}), 0)" if due else "0")
    )

    sql = f"""
        SELECT
            COUNT(*) AS invoices,
            {distinct_count("Advertiser Name")} AS advertisers,
            {distinct_count("Brand Name")} AS brands,
            {distinct_count("Team Responsible")} AS teams,
            {distinct_count("Accounts Manager")} AS managers,
            {distinct_count("Currency")} AS currencies,
            {coalesce_sum("Invoice amount")} AS invoice_amount,
            {coalesce_avg("Invoice amount")} AS avg_invoice,
            {coalesce_sum("Total Amount Due")} AS amount_due,
            {coalesce_sum("SGST")} AS sgst,
            {coalesce_sum("CGST")} AS cgst,
            {coalesce_sum("IGST")} AS igst,
            {coalesce_sum("TDS")} AS tds,
            {coalesce_sum("Bank Charges")} AS bank_charges,
            {coalesce_sum("Cheque Amt")} AS cheque_amount,
            {coalesce_sum("USD $ Billing value")} AS usd_billing,
            {fx_num} AS fx_weighted_num,
            {fx_den} AS fx_weighted_den,
            {risk_count} AS risky_count,
            {risk_due} AS risky_due,
            {force_count} AS force_count,
            {force_amt} AS force_amount,
            {intl_count} AS international_count,
            {diamond_count} AS diamond_count,
            {filled("Bank Realization date")} AS realized_count,
            {open_due_expr} AS open_due_count,
            {unrealized_due_expr} AS unrealized_due
        FROM invoices
        {where}
    """
    row = con.execute(sql, params).fetchone()
    invoices = int(row["invoices"] or 0)
    billed_v = float(row["invoice_amount"] or 0)
    due_v = float(row["amount_due"] or 0)
    fx_den_v = float(row["fx_weighted_den"] or 0)
    realized = int(row["realized_count"] or 0)
    sgst = float(row["sgst"] or 0)
    cgst = float(row["cgst"] or 0)
    igst = float(row["igst"] or 0)
    payload = {
        "invoices": invoices,
        "advertisers": int(row["advertisers"] or 0),
        "brands": int(row["brands"] or 0),
        "teams": int(row["teams"] or 0),
        "managers": int(row["managers"] or 0),
        "currencies": int(row["currencies"] or 0),
        "invoiceAmount": billed_v,
        "avgInvoice": float(row["avg_invoice"] or 0),
        "amountDue": due_v,
        "gst": sgst + cgst + igst,
        "sgst": sgst,
        "cgst": cgst,
        "igst": igst,
        "tds": float(row["tds"] or 0),
        "bankCharges": float(row["bank_charges"] or 0),
        "chequeAmount": float(row["cheque_amount"] or 0),
        "usdBilling": float(row["usd_billing"] or 0),
        "weightedFx": (float(row["fx_weighted_num"] or 0) / fx_den_v) if fx_den_v else 0,
        "riskyCount": int(row["risky_count"] or 0),
        "riskyDue": float(row["risky_due"] or 0),
        "forceCount": int(row["force_count"] or 0),
        "forceAmount": float(row["force_amount"] or 0),
        "internationalCount": int(row["international_count"] or 0),
        "diamondCount": int(row["diamond_count"] or 0),
        "realizedCount": realized,
        "unrealizedCount": max(invoices - realized, 0),
        "collectedEst": max(billed_v - due_v, 0.0),
        "openDueCount": int(row["open_due_count"] or 0),
        "unrealizedDue": float(row["unrealized_due"] or 0),
        "topAdvertisers": _top_group(con, label_to_sql, "Advertiser Name", "Invoice amount", where, params, 8),
        "topBrands": _top_group(con, label_to_sql, "Brand Name", "Invoice amount", where, params, 8),
        "topTeams": _top_group(con, label_to_sql, "Team Responsible", "Invoice amount", where, params, 8),
        "topManagers": _top_group(con, label_to_sql, "Accounts Manager", "Invoice amount", where, params, 8),
        "currencyMix": _top_group(con, label_to_sql, "Currency", "Invoice amount", where, params, 8),
        "topOutstanding": _top_group(con, label_to_sql, "Advertiser Name", "Total Amount Due", where, params, 8),
        "openInvoices": _open_invoices(con, label_to_sql, where, params, 10),
    }
    return _finalize(payload, scope="all")


def _top_group(con, label_to_sql, dim_label, metric_label, where, params, limit=5):
    dim = label_to_sql.get(dim_label)
    metric = label_to_sql.get(metric_label)
    if not dim:
        return []
    dim_q = '"' + str(dim).replace('"', "") + '"'
    metric_expr = (
        f'COALESCE(SUM("{str(metric).replace(chr(34), "")}"), 0)' if metric else "COUNT(*)"
    )
    sql = f"""
        SELECT
            CASE WHEN {dim_q} IS NULL OR TRIM(CAST({dim_q} AS TEXT)) IN ('', 'None')
                 THEN '(blank)' ELSE CAST({dim_q} AS TEXT) END AS name,
            {metric_expr} AS amount,
            COUNT(*) AS invoices
        FROM invoices
        {where}
        GROUP BY 1
        ORDER BY amount DESC, invoices DESC
        LIMIT {int(limit)}
    """
    rows = con.execute(sql, params).fetchall()
    return [
        {
            "name": str(r["name"]),
            "amount": round(float(r["amount"] or 0), 2),
            "invoices": int(r["invoices"] or 0),
        }
        for r in rows
    ]


def _top_from_rows(rows, key, limit=5, amount_key="Invoice amount"):
    buckets = {}
    for r in rows:
        name = str(r.get(key) or "").strip() or "(blank)"
        amt = r.get(amount_key)
        try:
            amt = float(amt or 0)
        except (TypeError, ValueError):
            amt = 0.0
        rec = buckets.setdefault(name, {"name": name, "amount": 0.0, "invoices": 0})
        rec["amount"] += amt
        rec["invoices"] += 1
    out = sorted(buckets.values(), key=lambda x: (-x["amount"], -x["invoices"], x["name"]))[:limit]
    for rec in out:
        rec["amount"] = round(rec["amount"], 2)
    return out


def _open_from_rows(rows, limit=10):
    def num(row, key):
        try:
            return float(row.get(key) or 0)
        except (TypeError, ValueError):
            return 0.0

    items = []
    for r in rows:
        due = num(r, "Total Amount Due")
        if due <= 0.009:
            continue
        items.append(
            {
                "invoice": str(r.get("Invoice Number") or "").strip(),
                "advertiser": str(r.get("Advertiser Name") or "").strip(),
                "brand": str(r.get("Brand Name") or "").strip(),
                "team": str(r.get("Team Responsible") or "").strip(),
                "due": round(due, 2),
                "billed": round(num(r, "Invoice amount"), 2),
            }
        )
    items.sort(key=lambda x: -x["due"])
    return items[:limit]


def _open_invoices(con, label_to_sql, where, params, limit=10):
    inv = _qid(label_to_sql, "Invoice Number")
    adv = _qid(label_to_sql, "Advertiser Name")
    brand = _qid(label_to_sql, "Brand Name")
    team = _qid(label_to_sql, "Team Responsible")
    due = _qid(label_to_sql, "Total Amount Due")
    billed = _qid(label_to_sql, "Invoice amount")
    if not due:
        return []
    extra = f"COALESCE({due}, 0) > 0.009"
    clause = f"{where} AND {extra}" if where else f"WHERE {extra}"
    inv_sql = f"CAST({inv} AS TEXT)" if inv else "NULL"
    adv_sql = f"CAST({adv} AS TEXT)" if adv else "NULL"
    brand_sql = f"CAST({brand} AS TEXT)" if brand else "NULL"
    team_sql = f"CAST({team} AS TEXT)" if team else "NULL"
    billed_sql = f"COALESCE({billed}, 0)" if billed else "0"
    sql = f"""
        SELECT
            {inv_sql} AS invoice,
            {adv_sql} AS advertiser,
            {brand_sql} AS brand,
            {team_sql} AS team,
            COALESCE({due}, 0) AS due,
            {billed_sql} AS billed
        FROM invoices
        {clause}
        ORDER BY due DESC
        LIMIT {int(limit)}
    """
    rows = con.execute(sql, params).fetchall()
    return [
        {
            "invoice": str(r["invoice"] or "").strip(),
            "advertiser": str(r["advertiser"] or "").strip(),
            "brand": str(r["brand"] or "").strip(),
            "team": str(r["team"] or "").strip(),
            "due": round(float(r["due"] or 0), 2),
            "billed": round(float(r["billed"] or 0), 2),
        }
        for r in rows
    ]


def _finalize(raw: dict, scope: str) -> dict:
    billed = float(raw.get("invoiceAmount") or 0)
    due = float(raw.get("amountDue") or 0)
    cheque = float(raw.get("chequeAmount") or 0)
    tds = float(raw.get("tds") or 0)
    gst = float(raw.get("gst") or 0)
    invoices = int(raw.get("invoices") or 0)
    realized = int(raw.get("realizedCount") or 0)
    risky_n = int(raw.get("riskyCount") or 0)
    collected = float(raw.get("collectedEst") or 0)

    def ratio(n, d):
        if not d:
            return 0.0
        return round(100.0 * float(n) / float(d), 1)

    raw.update(
        {
            "scope": scope,
            "invoiceAmount": round(billed, 2),
            "avgInvoice": round(float(raw.get("avgInvoice") or 0), 2),
            "amountDue": round(due, 2),
            "gst": round(gst, 2),
            "sgst": round(float(raw.get("sgst") or 0), 2),
            "cgst": round(float(raw.get("cgst") or 0), 2),
            "igst": round(float(raw.get("igst") or 0), 2),
            "tds": round(tds, 2),
            "bankCharges": round(float(raw.get("bankCharges") or 0), 2),
            "chequeAmount": round(cheque, 2),
            "usdBilling": round(float(raw.get("usdBilling") or 0), 2),
            "weightedFx": round(float(raw.get("weightedFx") or 0), 4),
            "riskyDue": round(float(raw.get("riskyDue") or 0), 2),
            "forceAmount": round(float(raw.get("forceAmount") or 0), 2),
            "collectedEst": round(collected, 2),
            "netAfterTds": round(billed - tds, 2),
            "outstandingPct": ratio(due, billed),
            "collectionPct": ratio(collected, billed),
            "chequeCoveragePct": ratio(cheque, billed),
            "tdsPct": ratio(tds, billed),
            "gstPct": round((gst / billed) * 100, 1) if billed else 0.0,
            "taxLiability": round(gst + tds, 2),
            "realizedPct": ratio(realized, invoices),
            "riskyPct": ratio(risky_n, invoices),
            "internationalPct": ratio(raw.get("internationalCount") or 0, invoices),
            "diamondPct": ratio(raw.get("diamondCount") or 0, invoices),
            "avgDue": round(due / invoices, 2) if invoices else 0,
            "openDueCount": int(raw.get("openDueCount") or 0),
            "unrealizedDue": round(float(raw.get("unrealizedDue") or 0), 2),
            "topOutstanding": raw.get("topOutstanding") or [],
            "topBrands": raw.get("topBrands") or [],
            "openInvoices": raw.get("openInvoices") or [],
        }
    )
    return raw


def _clean_ids(values: Sequence[str]) -> list:
    out = []
    seen = set()
    for raw in values or []:
        text = str(raw or "").strip()[:80]
        if not text or text in seen:
            continue
        seen.add(text)
        out.append(text)
        if len(out) >= 500:
            break
    return out


def _is_yes(value: str) -> bool:
    return value.lower() in YES_TOKENS


def _qid(label_to_sql: dict, label: str) -> Optional[str]:
    sql = label_to_sql.get(label)
    if not sql:
        return None
    return '"' + str(sql).replace('"', "") + '"'


def list_filter_options(db_path: str) -> dict:
    con = open_db(db_path)
    try:
        meta = {r[0]: r[1] for r in con.execute("SELECT key, value FROM meta")}
        _, _, label_to_sql = _schema_from_meta(meta)

        def distinct(label, limit=400):
            col = _qid(label_to_sql, label)
            if not col:
                return []
            rows = con.execute(
                f"""
                SELECT DISTINCT CAST({col} AS TEXT) AS v
                FROM invoices
                WHERE {col} IS NOT NULL AND TRIM(CAST({col} AS TEXT)) NOT IN ('', 'None')
                ORDER BY v
                LIMIT {int(limit) * 3}
                """
            ).fetchall()
            seen = {}
            for r in rows:
                raw = " ".join(str(r["v"] or "").split())
                key = raw.casefold()
                if key and key not in seen:
                    seen[key] = raw
                if len(seen) >= limit:
                    break
            return list(seen.values())

        inv = _qid(label_to_sql, "Invoice Date")
        date_min = date_max = ""
        if inv:
            row = con.execute(
                f"""
                SELECT MIN(dsr_date({inv})) AS dmin,
                       MAX(dsr_date({inv})) AS dmax
                FROM invoices
                WHERE dsr_date({inv}) != ''
                """
            ).fetchone()
            date_min = str(row["dmin"] or "")
            date_max = str(row["dmax"] or "")
        return {
            "brands": distinct("Brand Name"),
            "teams": distinct("Team Responsible"),
            "currencies": distinct("Currency", 40),
            "dateMin": date_min,
            "dateMax": date_max,
        }
    finally:
        con.close()


def compute_charts(db_path: str, search: str = "", filters: Optional[dict] = None) -> dict:
    con = open_db(db_path)
    try:
        meta = {r[0]: r[1] for r in con.execute("SELECT key, value FROM meta")}
        labels, sql_cols, label_to_sql = _schema_from_meta(meta)
        where, params = compose_where(search, sql_cols, label_to_sql, filters)
        billed = _qid(label_to_sql, "Invoice amount") or "0"
        due_col = _qid(label_to_sql, "Total Amount Due") or "0"
        date_field = (filters or {}).get("dateField") or "Invoice Date"
        if date_field not in ("Invoice Date", "Service Period"):
            date_field = "Invoice Date"
        date_col = _qid(label_to_sql, date_field) or _qid(label_to_sql, "Invoice Date") or _qid(label_to_sql, "Service Period")
        trend = []
        if date_col and billed != "0":
            rows = con.execute(
                f"""
                SELECT dsr_date({date_col}) AS bucket,
                       COALESCE(SUM({billed}), 0) AS amount,
                       COALESCE(SUM({due_col}), 0) AS due,
                       COUNT(*) AS invoices
                FROM invoices
                {where}
                GROUP BY 1
                HAVING bucket IS NOT NULL AND bucket != ''
                ORDER BY bucket
                LIMIT 180
                """,
                params,
            ).fetchall()
            trend = [
                {
                    "date": r["bucket"],
                    "amount": round(float(r["amount"] or 0), 2),
                    "due": round(float(r["due"] or 0), 2),
                    "invoices": int(r["invoices"] or 0),
                }
                for r in rows
            ]

        advertisers = _top_group(con, label_to_sql, "Advertiser Name", "Invoice amount", where, params, 10)
        brands = _top_group(con, label_to_sql, "Brand Name", "Invoice amount", where, params, 10)
        banks = _top_group(con, label_to_sql, "Bank", "Cheque Amt", where, params, 8)

        tax = con.execute(
            f"""
            SELECT
                COALESCE(SUM({_qid(label_to_sql, 'SGST') or '0'}),0) AS sgst,
                COALESCE(SUM({_qid(label_to_sql, 'CGST') or '0'}),0) AS cgst,
                COALESCE(SUM({_qid(label_to_sql, 'IGST') or '0'}),0) AS igst,
                COALESCE(SUM({_qid(label_to_sql, 'TDS') or '0'}),0) AS tds,
                COALESCE(SUM({_qid(label_to_sql, 'Bank Charges') or '0'}),0) AS charges
            FROM invoices {where}
            """,
            params,
        ).fetchone()

        team_col = _qid(label_to_sql, "Team Responsible")
        am_col = _qid(label_to_sql, "Accounts Manager")
        leaderboard = {"teams": [], "series": []}
        if team_col and am_col and billed != "0":
            top_teams = [r["name"] for r in _top_group(con, label_to_sql, "Team Responsible", "Invoice amount", where, params, 8)]
            if top_teams:
                placeholders = ",".join(["?"] * len(top_teams))
                extra = f"CAST({team_col} AS TEXT) IN ({placeholders})"
                tw = f"{where} AND {extra}" if where else f"WHERE {extra}"
                tparams = [*params, *top_teams]
                stacked = con.execute(
                    f"""
                    SELECT CAST({team_col} AS TEXT) AS team,
                           CASE WHEN {am_col} IS NULL OR TRIM(CAST({am_col} AS TEXT)) IN ('','None')
                                THEN '(unassigned)' ELSE CAST({am_col} AS TEXT) END AS manager,
                           COALESCE(SUM({billed}),0) AS amount
                    FROM invoices
                    {tw}
                    GROUP BY 1, 2
                    ORDER BY amount DESC
                    LIMIT 80
                    """,
                    tparams,
                ).fetchall()
                managers = []
                seen = set()
                for r in stacked:
                    name = str(r["manager"])
                    if name not in seen:
                        seen.add(name)
                        managers.append(name)
                    if len(managers) >= 6:
                        break
                series = []
                for mgr in managers:
                    series.append({
                        "name": mgr,
                        "data": [
                            round(sum(float(r["amount"] or 0) for r in stacked if str(r["team"]) == team and str(r["manager"]) == mgr), 2)
                            for team in top_teams
                        ],
                    })
                leaderboard = {"teams": top_teams, "series": series}

        risk = _qid(label_to_sql, "Force Billings and Risky Debtors")
        realized_col = _qid(label_to_sql, "Bank Realization date")
        realized_expr = (
            f"SUM(CASE WHEN {realized_col} IS NOT NULL AND TRIM(CAST({realized_col} AS TEXT)) NOT IN ('','None') THEN 1 ELSE 0 END)"
            if realized_col else "0"
        )
        unrealized_expr = (
            f"SUM(CASE WHEN {realized_col} IS NULL OR TRIM(CAST({realized_col} AS TEXT)) IN ('','None') THEN 1 ELSE 0 END)"
            if realized_col else "COUNT(*)"
        )
        risky_expr = (
            f"SUM(CASE WHEN LOWER(CAST({risk} AS TEXT)) LIKE '%risk%' THEN 1 ELSE 0 END)"
            if risk else "0"
        )
        row = con.execute(
            f"""
            SELECT {realized_expr} AS realized,
                   {unrealized_expr} AS unrealized,
                   {risky_expr} AS risky
            FROM invoices {where}
            """,
            params,
        ).fetchone()
        funnel = {
            "realized": int(row["realized"] or 0),
            "unrealized": int(row["unrealized"] or 0),
            "risky": int(row["risky"] or 0),
            "bankCharges": round(float(tax["charges"] or 0), 2),
        }

        return {
            "trend": trend,
            "advertisers": advertisers,
            "brands": brands,
            "tax": {
                "SGST": round(float(tax["sgst"] or 0), 2),
                "CGST": round(float(tax["cgst"] or 0), 2),
                "IGST": round(float(tax["igst"] or 0), 2),
                "TDS": round(float(tax["tds"] or 0), 2),
            },
            "leaderboard": leaderboard,
            "funnel": funnel,
            "banks": banks,
        }
    finally:
        con.close()


def iter_export_rows(db_path: str, search: str = "", filters: Optional[dict] = None, limit: int = 75000):
    con = open_db(db_path)
    try:
        meta = {r[0]: r[1] for r in con.execute("SELECT key, value FROM meta")}
        labels, sql_cols, label_to_sql = _schema_from_meta(meta)
        where, params = compose_where(search, sql_cols, label_to_sql, filters)
        col_sql = ", ".join(f'"{c}"' for c in sql_cols)
        rows = con.execute(
            f"SELECT {col_sql} FROM invoices {where} LIMIT ?",
            [*params, int(limit)],
        )
        yield labels
        for raw in rows:
            rec = []
            for sql in sql_cols:
                val = raw[sql]
                rec.append("" if val is None else val)
            yield rec
    finally:
        con.close()


def compute_team_performance(
    db_path: str,
    search: str = "",
    filters: Optional[dict] = None,
    month: str = "all",
    segment: str = "all",
    aging: str = "all",
    group_by: str = "manager",
) -> dict:
    """Team / collector performance from Daily Sales Register fields only."""
    filters = filters or {}
    con = open_db(db_path)
    try:
        meta = {r[0]: r[1] for r in con.execute("SELECT key, value FROM meta")}
        labels, sql_cols, label_to_sql = _schema_from_meta(meta)
        where, params = compose_where(search, sql_cols, label_to_sql, filters)
        extra, extra_params = _team_extra_where(label_to_sql, month, segment, aging)
        where, params = _merge_where(where, params, extra, extra_params)

        billed = _qid(label_to_sql, "Invoice amount") or "0"
        due = _qid(label_to_sql, "Total Amount Due") or "0"
        cheque = _qid(label_to_sql, "Cheque Amt") or "0"
        adv = _qid(label_to_sql, "Advertiser Name")
        team = _qid(label_to_sql, "Team Responsible")
        mgr = _qid(label_to_sql, "Accounts Manager")
        inv_date = _qid(label_to_sql, "Invoice Date")
        bank = _qid(label_to_sql, "Bank Realization date")
        flag = _qid(label_to_sql, "Force Billings and Risky Debtors")
        intl = _qid(label_to_sql, "International client")

        inv_iso = f"dsr_date({inv_date})" if inv_date else "''"
        bank_iso = f"dsr_date({bank})" if bank else "''"
        age_expr = (
            f"CAST(julianday('now') - julianday({inv_iso}) AS INTEGER)"
            if inv_date
            else "NULL"
        )
        dso_expr = (
            f"AVG(CASE WHEN {inv_iso} != '' AND {bank_iso} != '' "
            f"THEN julianday({bank_iso}) - julianday({inv_iso}) END)"
            if inv_date and bank
            else "NULL"
        )
        realized_expr = (
            f"SUM(CASE WHEN {bank_iso} != '' THEN 1 ELSE 0 END)" if bank else "0"
        )
        risky_expr = (
            f"SUM(CASE WHEN LOWER(CAST({flag} AS TEXT)) LIKE '%risk%' THEN 1 ELSE 0 END"
            ")"
            if flag
            else "0"
        )
        force_expr = (
            f"SUM(CASE WHEN LOWER(CAST({flag} AS TEXT)) LIKE '%force%' THEN 1 ELSE 0 END"
            ")"
            if flag
            else "0"
        )
        intl_expr = (
            f"SUM(CASE WHEN LOWER(TRIM(CAST({intl} AS TEXT))) "
            f"IN ('yes','y','true','1','international') THEN 1 ELSE 0 END)"
            if intl
            else "0"
        )
        open_expr = f"SUM(CASE WHEN COALESCE({due}, 0) > 0.009 THEN 1 ELSE 0 END)" if due != "0" else "0"
        accounts_expr = (
            f"COUNT(DISTINCT CASE WHEN {adv} IS NULL OR TRIM(CAST({adv} AS TEXT)) IN ('','None') "
            f"THEN NULL ELSE {adv} END)"
            if adv
            else "0"
        )

        totals = con.execute(
            f"""
            SELECT
                COUNT(*) AS invoices,
                {accounts_expr} AS accounts,
                COALESCE(SUM({billed}), 0) AS billed,
                COALESCE(SUM({due}), 0) AS due,
                COALESCE(SUM({cheque}), 0) AS collected,
                {realized_expr} AS realized,
                {risky_expr} AS risky,
                {force_expr} AS force_n,
                {intl_expr} AS international,
                {open_expr} AS open_invoices,
                {dso_expr} AS avg_dso
            FROM invoices
            {where}
            """,
            params,
        ).fetchone()

        dim = mgr if group_by != "team" else team
        dim_label = "Accounts Manager" if group_by != "team" else "Team Responsible"
        other = team if group_by != "team" else mgr
        other_label = "team" if group_by != "team" else "manager"
        people = []
        if dim:
            dim_name = (
                f"CASE WHEN {dim} IS NULL OR TRIM(CAST({dim} AS TEXT)) IN ('','None') "
                f"THEN '(unassigned)' ELSE TRIM(CAST({dim} AS TEXT)) END"
            )
            other_name = (
                f"CASE WHEN {other} IS NULL OR TRIM(CAST({other} AS TEXT)) IN ('','None') "
                f"THEN '' ELSE TRIM(CAST({other} AS TEXT)) END"
                if other
                else "''"
            )
            rows = con.execute(
                f"""
                SELECT
                    {dim_name} AS collector,
                    MAX({other_name}) AS secondary,
                    {accounts_expr} AS accounts,
                    COUNT(*) AS invoices,
                    COALESCE(SUM({billed}), 0) AS billed,
                    COALESCE(SUM({due}), 0) AS ar_balance,
                    COALESCE(SUM({cheque}), 0) AS collected,
                    {realized_expr} AS realized,
                    {risky_expr} AS risky,
                    {open_expr} AS open_invoices,
                    {dso_expr} AS avg_dso
                FROM invoices
                {where}
                GROUP BY 1
                ORDER BY collected DESC, ar_balance DESC, invoices DESC
                LIMIT 250
                """,
                params,
            ).fetchall()
            for r in rows:
                billed_v = float(r["billed"] or 0)
                collected_v = float(r["collected"] or 0)
                due_v = float(r["ar_balance"] or 0)
                invoices = int(r["invoices"] or 0)
                people.append(
                    {
                        "collector": str(r["collector"] or "(unassigned)"),
                        other_label: str(r["secondary"] or ""),
                        "accounts": int(r["accounts"] or 0),
                        "invoices": invoices,
                        "billed": round(billed_v, 2),
                        "arBalance": round(due_v, 2),
                        "collected": round(collected_v, 2),
                        "collectionRate": round(100.0 * collected_v / billed_v, 2) if billed_v else 0.0,
                        "realized": int(r["realized"] or 0),
                        "risky": int(r["risky"] or 0),
                        "openInvoices": int(r["open_invoices"] or 0),
                        "avgDso": round(float(r["avg_dso"] or 0), 1) if r["avg_dso"] is not None else None,
                    }
                )

        months = []
        if inv_date:
            month_rows = con.execute(
                f"""
                SELECT DISTINCT substr({inv_iso}, 1, 7) AS ym
                FROM invoices
                WHERE {inv_iso} != ''
                ORDER BY ym DESC
                LIMIT 24
                """
            ).fetchall()
            months = [str(r["ym"]) for r in month_rows if r["ym"]]

        billed_v = float(totals["billed"] or 0)
        collected_v = float(totals["collected"] or 0)
        due_v = float(totals["due"] or 0)
        invoices = int(totals["invoices"] or 0)
        return {
            "ok": True,
            "groupBy": "team" if group_by == "team" else "manager",
            "dimLabel": dim_label,
            "months": months,
            "summary": {
                "invoices": invoices,
                "accounts": int(totals["accounts"] or 0),
                "collectors": len(people),
                "billed": round(billed_v, 2),
                "collected": round(collected_v, 2),
                "arBalance": round(due_v, 2),
                "collectionRate": round(100.0 * collected_v / billed_v, 2) if billed_v else 0.0,
                "avgDso": round(float(totals["avg_dso"] or 0), 1) if totals["avg_dso"] is not None else None,
                "realized": int(totals["realized"] or 0),
                "openInvoices": int(totals["open_invoices"] or 0),
                "risky": int(totals["risky"] or 0),
                "riskyRate": round(100.0 * int(totals["risky"] or 0) / invoices, 2) if invoices else 0.0,
                "force": int(totals["force_n"] or 0),
                "international": int(totals["international"] or 0),
            },
            "people": people,
        }
    finally:
        con.close()


def _team_extra_where(label_to_sql: dict, month: str, segment: str, aging: str) -> tuple:
    clauses = []
    params = []
    inv = _qid(label_to_sql, "Invoice Date")
    due = _qid(label_to_sql, "Total Amount Due")
    flag = _qid(label_to_sql, "Force Billings and Risky Debtors")
    intl = _qid(label_to_sql, "International client")
    diamond = _qid(label_to_sql, "Diamond Seller")

    month = str(month or "all").strip()
    if month and month.lower() != "all" and inv:
        if month.lower() == "current":
            clauses.append(f"substr(dsr_date({inv}), 1, 7) = substr(date('now'), 1, 7)")
        elif len(month) >= 7 and month[4] == "-":
            ym = month[:7]
            clauses.append(f"substr(dsr_date({inv}), 1, 7) = ?")
            params.append(ym)

    segment = str(segment or "all").strip().lower()
    if segment == "international" and intl:
        clauses.append(
            f"LOWER(TRIM(CAST({intl} AS TEXT))) IN ('yes','y','true','1','international')"
        )
    elif segment == "domestic" and intl:
        clauses.append(
            f"LOWER(TRIM(CAST({intl} AS TEXT))) NOT IN ('yes','y','true','1','international')"
        )
    elif segment == "diamond" and diamond:
        clauses.append(
            f"LOWER(TRIM(CAST({diamond} AS TEXT))) IN ('yes','y','true','1','diamond')"
        )
    elif segment == "risky" and flag:
        clauses.append(f"LOWER(CAST({flag} AS TEXT)) LIKE '%risk%'")
    elif segment == "force" and flag:
        clauses.append(f"LOWER(CAST({flag} AS TEXT)) LIKE '%force%'")

    aging = str(aging or "all").strip().lower()
    if aging not in ("", "all") and inv and due:
        age = f"CAST(julianday('now') - julianday(dsr_date({inv})) AS INTEGER)"
        open_due = f"COALESCE({due}, 0) > 0.009 AND dsr_date({inv}) != ''"
        if aging in ("0-30", "current"):
            clauses.append(f"{open_due} AND {age} BETWEEN 0 AND 30")
        elif aging in ("31-60", "30-60"):
            clauses.append(f"{open_due} AND {age} BETWEEN 31 AND 60")
        elif aging in ("61-90", "60-90"):
            clauses.append(f"{open_due} AND {age} BETWEEN 61 AND 90")
        elif aging in ("90+", "90"):
            clauses.append(f"{open_due} AND {age} >= 90")

    if not clauses:
        return "", []
    return "WHERE " + " AND ".join(clauses), params
