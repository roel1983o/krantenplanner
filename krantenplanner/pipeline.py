import uuid
from pathlib import Path

from .def1_kordiam import run_def1
from .def2_planner import run_def2
from .def3_pdf import run_def3

ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"

def run_pipeline(*, kordiam_report_xlsx: str, posities_xlsx: str, workdir: str) -> dict:
    wd = Path(workdir)
    wd.mkdir(parents=True, exist_ok=True)

    mappingregels = str(ASSETS_DIR / "Mappingregels parser.xlsx")
    templates = str(ASSETS_DIR / "Templates.xlsx")
    beslispad_spread = str(ASSETS_DIR / "Beslispad Spread.xlsx")
    beslispad_ep = str(ASSETS_DIR / "Beslispad EP.xlsx")
    hoe_vaak = str(ASSETS_DIR / "Hoe vaak komt wat voor.xlsx")
    template_dir = str(ASSETS_DIR / "template_jpgs")

    # DEF1 output
    verhalen_out = str((wd / "Verhalenaanbod.xlsx").resolve())
    verhalenaanbod_xlsx = run_def1(kordiam_report_xlsx, mappingregels, verhalen_out)

    # DEF2 output
    krantenplanning_xlsx = str((wd / "Krantenplanning.xlsx").resolve())
    run_def2(
        templates_path=templates,
        beslispad_spread_path=beslispad_spread,
        beslispad_ep_path=beslispad_ep,
        posities_path=posities_xlsx,
        verhalenaanbod_path=verhalenaanbod_xlsx,
        out_path=krantenplanning_xlsx,
    )

    # DEF3 output
    handout_pdf = str((wd / "handout_modern_v3.pdf").resolve())
    run_def3(
        planning_xlsx=krantenplanning_xlsx,
        mapping_xlsx=hoe_vaak,
        template_dir=template_dir,
        out_pdf=handout_pdf,
    )

    return {
        "verhalenaanbod_xlsx": verhalenaanbod_xlsx,
        "krantenplanning_xlsx": krantenplanning_xlsx,
        "handout_pdf": handout_pdf,
    }
