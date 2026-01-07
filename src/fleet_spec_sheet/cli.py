import argparse
from pathlib import Path
from .generator import generate_spec_sheet

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--proof", required=True, type=Path)
    ap.add_argument("--print-zip", required=True, type=Path)
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--customer", default="Cornerstone Concrete")
    ap.add_argument("--job-no", default="")
    ap.add_argument("--unit-no", default="")
    ap.add_argument("--vehicle", default="")
    args = ap.parse_args()

    generate_spec_sheet(
        proof=args.proof,
        print_zip=args.print_zip,
        out_pdf=args.out,
        customer=args.customer,
        job_no=args.job_no,
        unit_no=args.unit_no,
        vehicle=args.vehicle,
    )

if __name__ == "__main__":
    main()
