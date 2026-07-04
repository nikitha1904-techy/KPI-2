from pathlib import Path
import shutil
import tempfile

from pipeline import SAMPLE_DIR, run_dynamic_pipeline, validate_reference_output

with tempfile.TemporaryDirectory() as tmp:
    out = Path(tmp)
    excel, _, meta = run_dynamic_pipeline(
        [SAMPLE_DIR / "report_a.json", SAMPLE_DIR / "report_b.json"],
        SAMPLE_DIR / "life_annuity_actuarial_reserving_kpi_ontology.xlsx",
        out,
    )
    print("mode:", meta.get("mode"))
    checks = validate_reference_output(excel)
    print("sheets_ok:", checks["sheets_ok"])
    for sheet, result in checks["row_checks"].items():
        print(sheet, result)
