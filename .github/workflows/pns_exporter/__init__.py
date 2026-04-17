"""
PnS NURBS Exporter — subprocess version
────────────────────────────────────────
This add-on exports the active Blender mesh as a NURBS STEP / IGES surface
by calling the standalone PolyhedralSplines.exe (shipped in  bin/  alongside
this file) as a subprocess.

No compiled Python extensions, no DLL hassles.  The .exe is statically linked
against the Visual C++ runtime so it runs on any Windows machine without
administrator rights or system installs.
"""

bl_info = {
    "name":        "PnS NURBS Exporter",
    "author":      "Based on SurfLab Polyhedral Net Splines (University of Florida)",
    "version":     (2, 0, 0),
    "blender":     (4, 0, 0),
    "location":    "File > Export > PnS NURBS",
    "description": "Export mesh as a NURBS surface (STEP / IGES) via PolyhedralSplines.exe",
    "category":    "Import-Export",
}

import os
import sys
import subprocess
import tempfile
import shutil
import traceback
import platform

import bpy
from bpy_extras.io_utils import ExportHelper
from bpy.props import StringProperty, BoolProperty, EnumProperty


# ─────────────────────────────────────────────────────────────────────────────
# Locate the bundled executable
# ─────────────────────────────────────────────────────────────────────────────

def _exe_path() -> str:
    """Return the expected path to PolyhedralSplines.exe inside the add-on."""
    name = "PolyhedralSplines.exe" if platform.system() == "Windows" else "PolyhedralSplines"
    return os.path.join(os.path.dirname(__file__), "bin", name)


def _exe_available() -> bool:
    return os.path.isfile(_exe_path())


# ─────────────────────────────────────────────────────────────────────────────
# Export mesh as OBJ (the format PolyhedralSplines.exe accepts)
# ─────────────────────────────────────────────────────────────────────────────

def _write_obj(obj, out_path: str, apply_modifiers: bool = True) -> None:
    """
    Write the given Blender object as an OBJ file that PolyhedralSplines.exe
    can read.  Writes only vertex positions and face indices — no UVs, no
    normals, no materials.  Subdivision Surface modifiers are always skipped
    so the cage mesh is used as the control net.

    This is a minimal OBJ writer so we don't depend on any particular
    Blender OBJ exporter being available (Blender 4.x renamed the built-in
    exporter a couple of times).
    """
    # Temporarily hide any Subdivision Surface modifiers
    saved = {}
    for m in obj.modifiers:
        if m.type == 'SUBSURF':
            saved[m.name] = m.show_viewport
            m.show_viewport = False

    try:
        if apply_modifiers:
            depsgraph = bpy.context.evaluated_depsgraph_get()
            eval_obj  = obj.evaluated_get(depsgraph)
            mesh      = eval_obj.to_mesh()
        else:
            eval_obj  = None
            mesh      = obj.data

        matrix = obj.matrix_world

        with open(out_path, "w", encoding="utf-8") as f:
            f.write("# Exported from Blender for PolyhedralSplines\n")
            for v in mesh.vertices:
                world = matrix @ v.co
                f.write("v {:.9f} {:.9f} {:.9f}\n".format(world.x, world.y, world.z))
            for p in mesh.polygons:
                # OBJ face indices are 1-based
                indices = " ".join(str(i + 1) for i in p.vertices)
                f.write("f {}\n".format(indices))

        if eval_obj is not None:
            eval_obj.to_mesh_clear()

    finally:
        # Restore SubD modifier visibility
        for m in obj.modifiers:
            if m.name in saved:
                m.show_viewport = saved[m.name]


# ─────────────────────────────────────────────────────────────────────────────
# Run PolyhedralSplines.exe
# ─────────────────────────────────────────────────────────────────────────────

def _run_pns(obj_path: str, out_path: str, fmt: str, degree_raise: bool) -> tuple[bool, str]:
    """
    Invoke PolyhedralSplines.exe.  The current SurfLab CLI writes its output
    file in the same folder as the input OBJ, using the input basename plus
    the chosen extension.  We run it in a temp folder, then move the
    generated file to the user's requested path.

    Returns (ok, message).
    """
    exe = _exe_path()

    # Command line arguments per SurfLab README:
    #   -f <format>   bv | igs  (default bv)
    #   -d            raise degree-2 patches to degree 3
    #   <input.obj>   positional
    args = [exe, "-f", {"STEP": "igs", "IGES": "igs", "BV": "bv"}[fmt]]
    if degree_raise:
        args.append("-d")
    args.append(obj_path)

    # STEP export: the .exe only emits BV and IGS natively, so we always
    # request IGS and rename.  (Most CAD apps read IGS just fine; STEP
    # export would require using the STEPWriter class from the C++ library,
    # which this subprocess approach doesn't expose.)
    #
    # If STEP is critical for the user, we can run through FreeCAD or
    # trimesh for IGS→STEP conversion in a future iteration.

    try:
        result = subprocess.run(
            args,
            capture_output=True,
            text=True,
            timeout=600,  # 10 min safety cap
        )
    except FileNotFoundError:
        return False, (
            "PolyhedralSplines.exe was not found at:\n  {}\n\n"
            "Re-install the add-on zip and make sure bin/ contains the .exe."
        ).format(exe)
    except subprocess.TimeoutExpired:
        return False, "Conversion timed out after 10 minutes."

    if result.returncode != 0:
        return False, (
            "PolyhedralSplines.exe exited with code {}.\n\n"
            "STDOUT:\n{}\n\nSTDERR:\n{}"
        ).format(result.returncode, result.stdout, result.stderr)

    # Find the file the exe produced.  It is written next to the input OBJ.
    tmp_dir   = os.path.dirname(obj_path)
    stem      = os.path.splitext(os.path.basename(obj_path))[0]
    ext_map   = {"STEP": ".igs", "IGES": ".igs", "BV": ".bv"}
    produced  = os.path.join(tmp_dir, stem + ext_map[fmt])

    if not os.path.isfile(produced):
        # Some builds of the exe use different naming — fall back to finding
        # anything newly created in the temp folder.
        candidates = [
            os.path.join(tmp_dir, fn)
            for fn in os.listdir(tmp_dir)
            if fn.endswith((".igs", ".bv", ".step"))
        ]
        if not candidates:
            return False, (
                "Conversion finished but no output file was found in {}."
                "\n\nSTDOUT:\n{}"
            ).format(tmp_dir, result.stdout)
        produced = candidates[0]

    # Move to user's chosen location
    try:
        shutil.move(produced, out_path)
    except Exception as e:
        return False, "Could not move output to {}: {}".format(out_path, e)

    return True, result.stdout.strip() or "Export successful."


# ─────────────────────────────────────────────────────────────────────────────
# Export operator
# ─────────────────────────────────────────────────────────────────────────────

class PNS_OT_Export(bpy.types.Operator, ExportHelper):
    """Export active mesh as a NURBS surface via PolyhedralSplines"""
    bl_idname  = "export_mesh.pns_nurbs"
    bl_label   = "Export PnS NURBS"
    bl_options = {'REGISTER'}

    filename_ext: StringProperty(default=".igs", options={'HIDDEN'})
    filter_glob:  StringProperty(default="*.igs;*.bv", options={'HIDDEN'})

    export_format: EnumProperty(
        name="Format",
        items=[
            ("IGES", "IGES (.igs)",
             "Universal CAD interchange format — opens in Fusion 360, SolidWorks, Rhino, FreeCAD"),
            ("BV",   "BV   (.bv)",
             "SurfLab BView format — for the BView viewer only"),
        ],
        default="IGES",
    )

    degree_raise: BoolProperty(
        name="Raise to Degree 3",
        description="Make all patches bi-cubic.  Recommended — produces a smoother surface.",
        default=True,
    )

    apply_modifiers: BoolProperty(
        name="Apply Modifiers",
        description="Bake Mirror, Bevel, etc. into the mesh before export. "
                    "Subdivision Surface is always excluded.",
        default=True,
    )

    def invoke(self, context, event):
        obj = context.active_object
        if obj:
            stem = bpy.path.clean_name(obj.name)
            ext  = {"IGES": ".igs", "BV": ".bv"}[self.export_format]
            base = os.path.dirname(bpy.data.filepath) or tempfile.gettempdir()
            self.filepath = os.path.join(base, stem + ext)
        context.window_manager.fileselect_add(self)
        return {'RUNNING_MODAL'}

    def draw(self, context):
        layout = self.layout
        layout.use_property_split = True
        layout.prop(self, "export_format")
        layout.prop(self, "degree_raise")
        layout.prop(self, "apply_modifiers")

    def execute(self, context):
        obj = context.active_object

        if obj is None or obj.type != 'MESH':
            self.report({'ERROR'}, "Select a mesh object before exporting.")
            return {'CANCELLED'}

        if not _exe_available():
            self.report({'ERROR'}, (
                "PolyhedralSplines.exe is missing.\n"
                "Expected at: {}\n\n"
                "Re-install the add-on zip."
            ).format(_exe_path()))
            return {'CANCELLED'}

        # Make sure the filepath ends with the right extension
        ext_map  = {"IGES": ".igs", "BV": ".bv"}
        filepath = bpy.path.ensure_ext(self.filepath, ext_map[self.export_format])

        # Create a clean scratch folder
        tmp_dir  = tempfile.mkdtemp(prefix="pns_")
        obj_path = os.path.join(tmp_dir, "input.obj")

        try:
            _write_obj(obj, obj_path, self.apply_modifiers)
        except Exception:
            shutil.rmtree(tmp_dir, ignore_errors=True)
            self.report({'ERROR'}, "Failed to write OBJ:\n" + traceback.format_exc())
            return {'CANCELLED'}

        ok, msg = _run_pns(obj_path, filepath, self.export_format, self.degree_raise)
        shutil.rmtree(tmp_dir, ignore_errors=True)

        if not ok:
            self.report({'ERROR'}, msg)
            return {'CANCELLED'}

        self.report({'INFO'}, "Exported to {}".format(os.path.basename(filepath)))
        return {'FINISHED'}


# ─────────────────────────────────────────────────────────────────────────────
# Verify operator — quick smoke test
# ─────────────────────────────────────────────────────────────────────────────

class PNS_OT_Verify(bpy.types.Operator):
    """Verify PolyhedralSplines.exe is available and working"""
    bl_idname = "pns.verify"
    bl_label  = "Verify PolyhedralSplines.exe"

    def execute(self, context):
        exe = _exe_path()

        if not os.path.isfile(exe):
            self.report({'ERROR'},
                "PolyhedralSplines.exe not found at:\n{}".format(exe))
            return {'CANCELLED'}

        # Try running it with --help; exit code 0 means it's loadable
        try:
            result = subprocess.run(
                [exe, "--help"],
                capture_output=True, text=True, timeout=10,
            )
        except Exception as e:
            self.report({'ERROR'}, "Could not run .exe: {}".format(e))
            return {'CANCELLED'}

        if result.returncode != 0:
            self.report({'ERROR'}, (
                ".exe ran but returned error code {}.\n"
                "If you see DLL errors, the .exe wasn't built with static "
                "runtime linking.\n\nStderr:\n{}"
            ).format(result.returncode, result.stderr))
            return {'CANCELLED'}

        self.report({'INFO'}, "PolyhedralSplines.exe is working.")
        return {'FINISHED'}


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar panel
# ─────────────────────────────────────────────────────────────────────────────

class PNS_PT_Panel(bpy.types.Panel):
    bl_label       = "PnS NURBS Export"
    bl_idname      = "PNS_PT_main"
    bl_space_type  = 'VIEW_3D'
    bl_region_type = 'UI'
    bl_category    = "PnS Export"

    def draw(self, context):
        layout = self.layout

        if _exe_available():
            layout.label(text="PolyhedralSplines.exe: ready", icon="CHECKMARK")
        else:
            row = layout.row()
            row.alert = True
            row.label(text="PolyhedralSplines.exe: missing", icon="ERROR")

        layout.separator()

        col = layout.column(align=True)
        col.scale_y = 1.5
        col.operator("export_mesh.pns_nurbs", text="Export Active Object", icon="EXPORT")
        col.operator("pns.verify", text="Verify .exe", icon="VIEWZOOM")


# ─────────────────────────────────────────────────────────────────────────────
# File > Export menu entry
# ─────────────────────────────────────────────────────────────────────────────

def _menu_func_export(self, context):
    self.layout.operator(PNS_OT_Export.bl_idname, text="PnS NURBS (.igs / .bv)")


# ─────────────────────────────────────────────────────────────────────────────
# Registration
# ─────────────────────────────────────────────────────────────────────────────

_classes = [PNS_OT_Export, PNS_OT_Verify, PNS_PT_Panel]


def register():
    for cls in _classes:
        bpy.utils.register_class(cls)
    bpy.types.TOPBAR_MT_file_export.append(_menu_func_export)


def unregister():
    bpy.types.TOPBAR_MT_file_export.remove(_menu_func_export)
    for cls in reversed(_classes):
        bpy.utils.unregister_class(cls)


if __name__ == "__main__":
    register()
