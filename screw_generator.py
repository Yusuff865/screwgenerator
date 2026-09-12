"""
SolidWorks Custom Screw Generator
----------------------------------
Connects to a running (or new) SolidWorks 2025 session via COM and builds a
simple screw part from user supplied dimensions. Threads must be added manually, which keeps
rebuild times sensible and is normally sufficient for drawings and BOMs.

Requirements:
    pip install pywin32
    SolidWorks 2025 installed and licensed on this machine (Windows only)

Usage:
    python screw_generator.py
"""

import win32com.client
import pythoncom


# ---------------------------------------------------------------------------
# Constants matching the SolidWorks API enumerations we use
# ---------------------------------------------------------------------------
swDocPART = 1
swThisConfiguration = 1
swAllConfiguration = 2
swSpecifyConfiguration = 3

swUnitsLinearMillimeters = 8  # for reference; new parts default to whatever
                               # template is used, so this isn't set directly


class ScrewParameters:
    """Holds every dimension the user can vary, all in millimetres."""

    def __init__(
        self,
        shank_diameter=6.0,
        shank_length=25.0,
        thread_pitch=1.0,
        head_diameter=10.0,
        head_height=4.0,
        head_style="hex",   # "hex" or "cylindrical"
        drive_style="hex",  # "hex", "phillips", "slot"
    ):
        self.shank_diameter = shank_diameter
        self.shank_length = shank_length
        self.thread_pitch = thread_pitch
        self.head_diameter = head_diameter
        self.head_height = head_height
        self.head_style = head_style
        self.drive_style = drive_style


def get_user_parameters():
    """Prompt on the console for each dimension, falling back to sensible
    defaults if the user just presses enter."""

    def ask(prompt, default):
        raw = input(f"{prompt} [{default}]: ").strip()
        return float(raw) if raw else default

    print("Enter screw dimensions in millimetres (press enter to accept default)\n")

    shank_diameter = ask("Shank diameter", 6.0)
    shank_length = ask("Shank length", 25.0)
    thread_pitch = ask("Thread pitch", 1.0)
    head_diameter = ask("Head diameter", 10.0)
    head_height = ask("Head height", 4.0)

    head_style = input("Head style, hex or cylindrical [hex]: ").strip().lower() or "hex"
    drive_style = input("Drive style, hex/phillips/slot [hex]: ").strip().lower() or "hex"

    return ScrewParameters(
        shank_diameter, shank_length, thread_pitch,
        head_diameter, head_height, head_style, drive_style,
    )


# A properly typed "Nothing" for the Callout argument SelectByID2 expects.
# Plain Python None fails with a 'Type mismatch' COM error under late
# binding, because COM cannot infer what type of null is meant. Wrapping it
# as an explicit VT_DISPATCH variant tells COM exactly what's intended,
# without needing gencache/makepy at all.
NOTHING_DISPATCH = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)


def connect_to_solidworks():
    """Attach to an already-open SolidWorks instance, or launch a new one."""
    try:
        sw_app = win32com.client.GetActiveObject("SldWorks.Application")
    except pythoncom.com_error:
        sw_app = win32com.client.Dispatch("SldWorks.Application")
        sw_app.Visible = True
    return sw_app


def new_part_document(sw_app):
    """Create a blank part from the default template."""
    template_path = sw_app.GetUserPreferenceStringValue(8)  # swDefaultTemplatePart
    model = sw_app.NewDocument(template_path, 0, 0, 0)
    if model is None:
        raise RuntimeError("SolidWorks did not return a new part document")
    return model


def build_shank(model, params):
    """Sketch a circle on the front plane and extrude it into the shank."""
    sketch_mgr = model.SketchManager
    model.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, NOTHING_DISPATCH, 0)
    sketch_mgr.InsertSketch(True)

    radius_m = (params.shank_diameter / 2.0) / 1000.0  # mm to metres
    sketch_mgr.CreateCircleByRadius(0, 0, 0, radius_m)

    sketch_mgr.InsertSketch(True)

    feature_mgr = model.FeatureManager
    depth_m = params.shank_length / 1000.0
    shank_feature = feature_mgr.FeatureExtrusion2(
        True, False, False, 0, 0, depth_m, 0.0,
        False, False, False, False, 0, 0,
        False, False, False, False, True, True, True, 0, 0, False,
    )
    return shank_feature


def build_head(model, params):
    """Sketch and extrude the head on top of the shank."""
    sketch_mgr = model.SketchManager
    model.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, NOTHING_DISPATCH, 0)
    sketch_mgr.InsertSketch(True)

    if params.head_style == "cylindrical":
        radius_m = (params.head_diameter / 2.0) / 1000.0
        sketch_mgr.CreateCircleByRadius(0, 0, 0, radius_m)
    else:
        # Hex head: six-sided polygon inscribed to the given diameter
        radius_m = (params.head_diameter / 2.0) / 1000.0
        sketch_mgr.CreatePolygon(0, 0, 0, radius_m, 0, 0, 6, False)

    sketch_mgr.InsertSketch(True)

    feature_mgr = model.FeatureManager
    depth_m = params.head_height / 1000.0
    head_feature = feature_mgr.FeatureExtrusion2(
        True, False, False, 0, 0, depth_m, 0.0,
        False, False, False, False, 0, 0,
        False, False, False, False, True, True, True, 0, 0, False,
    )
    return head_feature


def add_cosmetic_thread(model, params):
    """Adds a cosmetic thread annotation along the shank face rather than
    cutting a true helix, which keeps the model light."""
    feature_mgr = model.FeatureManager
    # InsertCosmeticThread2 expects the target face/edge to already be
    # selected; in a fuller implementation you would locate and select the
    # cylindrical shank face here before calling this.
    try:
        feature_mgr.InsertCosmeticThread2(
            0,                       # thread type: 0 = simple
            params.shank_diameter / 1000.0,
            params.shank_length / 1000.0,
            0,                       # end condition: blind
            0,
            False,
        )
    except Exception as exc:
        print(f"Could not add cosmetic thread automatically: {exc}")
        print("You may need to select the shank face manually first.")


def save_part(model, file_path):
    # Same root cause as the SelectByID2 issue earlier: the PDMComponents
    # and TexData arguments need an explicitly typed 'Nothing', not plain
    # Python None. The Errors/Warnings arguments have a similar problem -
    # they're ByRef longs, so a bare 0 doesn't carry enough type information
    # for COM either. Wrapping them as VT_BYREF|VT_I4 variants fixes both.
    errors_variant = win32com.client.VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_I4, 0
    )
    warnings_variant = win32com.client.VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_I4, 0
    )

    model.Extension.SaveAs3(
        file_path, 0, 0, NOTHING_DISPATCH, NOTHING_DISPATCH,
        errors_variant, warnings_variant,
    )

    return errors_variant.value, warnings_variant.value


def main():
    params = get_user_parameters()

    sw_app = connect_to_solidworks()
    model = new_part_document(sw_app)

    build_shank(model, params)
    build_head(model, params)
    add_cosmetic_thread(model, params)

    model.ForceRebuild3(False)

    output_path = input("\nSave part as (full path, .sldprt): ").strip()
    if output_path:
        errors, warnings = save_part(model, output_path)
        if errors:
            print(f"Save finished with {errors} error code(s)")
        else:
            print(f"Saved to {output_path}")


if __name__ == "__main__":
    main()