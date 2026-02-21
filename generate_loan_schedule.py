import datetime as dt
from datetime import date

try:
    import xlsxwriter
except ImportError as e:
    raise SystemExit("This script requires the 'xlsxwriter' package. Install with: pip install xlsxwriter")


def generate_workbook(file_path: str):
    # Defaults based on user's scenario
    loan_amount = 787_500.00
    start_date = dt.date(2022, 6, 22)
    original_term_months = 360

    # Rate periods: (Start Date, Annual Rate %, Term Months)
    # Term (months) is intended to be the *remaining amortization term starting from that period*.
    # Example for a 30y loan (360 months):
    # - Start: remaining term 360
    # - After 24 months: remaining term 336
    # - After 48 months: remaining term 312
    rate_periods = [
        (dt.date(2022, 6, 22), 0.0143, 360),
        (dt.date(2024, 6, 22), 0.0295, 336),
        (dt.date(2026, 6, 22), 0.0350, 312),  # user-editable
    ]

    # Prepayment events: (Date, Amount)
    prepayments = [
        (dt.date(2024, 6, 22), 50_000.00),
        (dt.date(2026, 6, 22), 50_000.00),
    ]

    wb = xlsxwriter.Workbook(file_path)

    # Formats
    fmt_title = wb.add_format({"bold": True, "font_size": 12})
    fmt_hdr = wb.add_format({"bold": True, "bg_color": "#F0F0F0", "border": 1})
    fmt_money = wb.add_format({"num_format": "$#,##0.00"})
    fmt_pct = wb.add_format({"num_format": "0.00%"})
    fmt_date = wb.add_format({"num_format": "yyyy-mm-dd"})
    fmt_int = wb.add_format({"num_format": "0"})
    fmt_text = wb.add_format({})

    # Inputs sheet
    sh_in = wb.add_worksheet("Inputs")
    sh_in.freeze_panes(2, 0)

    # General inputs
    sh_in.write(0, 0, "Loan inputs", fmt_title)
    sh_in.write(1, 0, "Loan Amount")
    sh_in.write_number(1, 1, loan_amount, fmt_money)

    sh_in.write(2, 0, "Start Date")
    sh_in.write_datetime(2, 1, dt.datetime.combine(start_date, dt.time()), fmt_date)

    sh_in.write(3, 0, "Original Term (months)")
    sh_in.write_number(3, 1, original_term_months, fmt_int)

    sh_in.write(5, 0, "Notes")
    sh_in.write(6, 0, "• In Rate changes, Term (months) should be the remaining term at that refinance/reset date.")
    sh_in.write(7, 0, "  E.g., a 30y (360 mo) loan after 24 payments has 336 months remaining.")
    sh_in.write(8, 0, "• Add more rate rows for future refinances (dates must be increasing).")
    sh_in.write(9, 0, "• Add prepayments by date; they reduce balance in that month.")

    # Rate changes table at columns D:F
    sh_in.write(1, 3, "Rate changes", fmt_title)
    sh_in.write(2, 3, "Start Date", fmt_hdr)
    sh_in.write(2, 4, "Annual Rate", fmt_hdr)
    sh_in.write(2, 5, "Term (months)", fmt_hdr)

    rate_start_row = 3
    for i, (sd, r, term) in enumerate(rate_periods):
        rrow = rate_start_row + i
        sh_in.write_datetime(rrow, 3, dt.datetime.combine(sd, dt.time()), fmt_date)
        sh_in.write_number(rrow, 4, r, fmt_pct)
        sh_in.write_number(rrow, 5, term, fmt_int)

    # Prepayments table at columns H:I
    sh_in.write(1, 7, "Prepayments", fmt_title)
    sh_in.write(2, 7, "Date", fmt_hdr)
    sh_in.write(2, 8, "Amount", fmt_hdr)

    prepay_start_row = 3
    for i, (pd, amt) in enumerate(prepayments):
        prow = prepay_start_row + i
        sh_in.write_datetime(prow, 7, dt.datetime.combine(pd, dt.time()), fmt_date)
        sh_in.write_number(prow, 8, amt, fmt_money)

    # Widen columns for readability
    sh_in.set_column(0, 0, 22)
    sh_in.set_column(1, 1, 18)
    sh_in.set_column(3, 3, 14)  # rate start date
    sh_in.set_column(4, 4, 12)  # rate
    sh_in.set_column(5, 5, 14)  # term
    sh_in.set_column(7, 7, 14)  # prepay date
    sh_in.set_column(8, 8, 14)  # prepay amt

    # Schedule sheet
    sh = wb.add_worksheet("Schedule")
    sh.freeze_panes(1, 0)

    headers = [
        "Period",
        "Date",
        "Annual Rate",
        "Monthly Rate",
        "Beginning Balance",
        "Is Rate Period Start",
        "Remaining Term (months)",
        "Payment",
        "Interest",
        "Scheduled Principal",
        "Prepayment",
        "Total Principal",
        "Ending Balance",
        "Cumulative Interest",
    ]
    for col, h in enumerate(headers):
        sh.write(0, col, h, fmt_hdr)

    # Column widths
    sh.set_column(0, 0, 8)
    sh.set_column(1, 1, 12)
    sh.set_column(2, 3, 12)
    sh.set_column(4, 4, 16)
    sh.set_column(5, 7, 18)
    sh.set_column(8, 13, 16)

    # Number of rows to generate (enough to cover full term)
    max_rows = 420  # 35 years coverage, safe for edits

    # Pre-calc absolute references for Inputs ranges
    # Rate lookup ranges: Inputs!D:D (start date), E:E (rate), F:F (term)
    rate_date_col = "D"
    rate_rate_col = "E"
    rate_term_col = "F"

    # Prepayment lookup ranges: Inputs!H:H (date), I:I (amount)
    prepay_date_col = "H"
    prepay_amt_col = "I"

    # Helper: Excel row index (1-based) from 0-based loop row
    def xl_row(n):
        return n + 1

    for i in range(max_rows):
        row = i + 2  # data starts at row 2 (1-based)
        prev_row = row - 1

        # Period number
        sh.write_number(i + 1, 0, i + 1, fmt_int)

        # Date = EDATE(Inputs!$B$3, Period-1)
        sh.write_formula(i + 1, 1, f"=EDATE(Inputs!$B$3,A{row}-1)", fmt_date)

        # Annual Rate = XLOOKUP(Date, Inputs!D:D, Inputs!E:E, , -1)
        sh.write_formula(i + 1, 2, f"=XLOOKUP(B{row},Inputs!{rate_date_col}:{rate_date_col},Inputs!{rate_rate_col}:{rate_rate_col},, -1)", fmt_pct)

        # Monthly Rate = Annual/12 (0 when balance is 0)
        sh.write_formula(i + 1, 3, f"=IF(E{row}<=0,0,C{row}/12)")

        # Beginning Balance: first row uses loan amount; else prior ending balance
        if i == 0:
            sh.write_formula(i + 1, 4, "=Inputs!$B$2", fmt_money)
        else:
            sh.write_formula(i + 1, 4, f"=M{prev_row}", fmt_money)

        # Is Rate Period Start: TRUE when date equals the matching period start
        sh.write_formula(i + 1, 5, f"=IF(B{row}=XLOOKUP(B{row},Inputs!{rate_date_col}:{rate_date_col},Inputs!{rate_date_col}:{rate_date_col},,-1),TRUE,FALSE)")

        # Remaining term (months):
        # - At the start of a rate period: pull remaining term from Inputs
        # - Otherwise: previous remaining term - 1
        if i == 0:
            sh.write_formula(
                i + 1,
                6,
                f"=XLOOKUP(B{row},Inputs!{rate_date_col}:{rate_date_col},Inputs!{rate_term_col}:{rate_term_col},,-1)",
                fmt_int,
            )
        else:
            sh.write_formula(
                i + 1,
                6,
                f"=IF(E{row}<=0,0,IF(F{row},XLOOKUP(B{row},Inputs!{rate_date_col}:{rate_date_col},Inputs!{rate_term_col}:{rate_term_col},,-1),MAX(0,G{prev_row}-1)))",
                fmt_int,
            )

        # Payment:
        # - Recalculate at the start of each rate period using remaining term (column G)
        # - Otherwise carry forward previous payment
        # - If remaining term is 0 or balance is 0 -> 0
        if i == 0:
            # First row has no prior payment, so always compute
            sh.write_formula(i + 1, 7, f"=IF(OR(E{row}<=0,G{row}<=0),0,-PMT(D{row},G{row},E{row}))", fmt_money)
        else:
            sh.write_formula(i + 1, 7, f"=IF(OR(E{row}<=0,G{row}<=0),0,IF(F{row},-PMT(D{row},G{row},E{row}),H{prev_row}))", fmt_money)


        # Interest = Beginning Balance * Monthly Rate
        sh.write_formula(i + 1, 8, f"=IF(E{row}<=0,0,E{row}*D{row})", fmt_money)

        # Scheduled Principal = min(balance, max(0, payment - interest))
        sh.write_formula(i + 1, 9, f"=IF(E{row}<=0,0,MIN(E{row},MAX(0,H{row}-I{row})))", fmt_money)

        # Prepayment = lookup amount for exact date; default 0
        sh.write_formula(i + 1, 10, f"=IFERROR(XLOOKUP(B{row},Inputs!{prepay_date_col}:{prepay_date_col},Inputs!{prepay_amt_col}:{prepay_amt_col},0,0),0)", fmt_money)


        # Total Principal (cap at balance)
        sh.write_formula(i + 1, 11, f"=IF(E{row}<=0,0,MIN(E{row},J{row}+K{row}))", fmt_money)

        # Ending Balance = BegBal - TotalPrincipal
        sh.write_formula(i + 1, 12, f"=MAX(0,E{row}-L{row})", fmt_money)

        # Cumulative Interest
        if i == 0:
            sh.write_formula(i + 1, 13, f"=I{row}", fmt_money)
        else:
            sh.write_formula(i + 1, 13, f"=N{prev_row}+I{row}", fmt_money)

    wb.close()


if __name__ == "__main__":
    # If the workbook is open in Excel, overwriting will fail on Windows.
    # Write a timestamped file by default to avoid PermissionError.
    ts = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    out_path = f"loan_schedule_{ts}.xlsx"
    generate_workbook(out_path)
    print(f"Created {out_path}")
