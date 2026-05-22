"""
ModuleExtractor: extract VBA module source code via the VBIDE COM interface.

Handles:
  - Standard modules, class modules, form code-behind, report code-behind
  - Password-protected VBA projects (graceful degradation: cataloged but no source)
  - Broken / missing references
"""

from __future__ import annotations

from typing import Any

from access_dissect.catalog.models import (
    ExtractionWarning,
    VBAModule,
    VBAModuleType,
    WarningSeverity,
)

# VBA module type constants from the VBE type library
VBA_MODULE_TYPE_STANDARD = 1    # vbext_ct_StdModule
VBA_MODULE_TYPE_CLASS = 2       # vbext_ct_ClassModule
VBA_MODULE_TYPE_FORM = 100      # vbext_ct_MSForm (used by Access form/report modules)
VBA_MODULE_TYPE_DOCUMENT = 100  # vbext_ct_Document (Access form/report code-behind)

# HRESULT for VBA project password protection
VBA_LOCKED_HRESULT = -2146827282  # 0x800A03EC


def _vba_module_type(component: Any) -> VBAModuleType:
    """Map a VBComponent.Type integer to VBAModuleType enum."""
    try:
        t = int(component.Type)
    except Exception:
        return VBAModuleType.STANDARD

    comp_name = ""
    try:
        comp_name = str(component.Name)
    except Exception:
        pass

    if comp_name.startswith("Form_"):
        return VBAModuleType.FORM_MODULE
    if comp_name.startswith("Report_"):
        return VBAModuleType.REPORT_MODULE
    if t == 2:
        return VBAModuleType.CLASS
    return VBAModuleType.STANDARD


def _parent_object_name(module_name: str, module_type: VBAModuleType) -> str | None:
    """Derive the parent form/report name from the module name convention."""
    if module_type == VBAModuleType.FORM_MODULE and module_name.startswith("Form_"):
        return module_name[len("Form_"):]
    if module_type == VBAModuleType.REPORT_MODULE and module_name.startswith("Report_"):
        return module_name[len("Report_"):]
    return None


def extract_all_modules(
    app: Any,
    vba_locked: bool = False,
) -> tuple[list[VBAModule], list[ExtractionWarning]]:
    """
    Extract all VBA modules from the VBE project.

    If vba_locked is True (already detected in PropertyExtractor), returns
    empty modules with is_vba_locked=True for each discovered component name.
    If vba_locked is False, attempts extraction and handles locking gracefully.
    """
    modules: list[VBAModule] = []
    warnings: list[ExtractionWarning] = []

    if app is None:
        return modules, warnings

    try:
        vbe = app.VBE
    except Exception as e:
        warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="VBE",
                warning_code="VBE_UNAVAILABLE",
                message=f"Could not access VBE: {e}",
                severity=WarningSeverity.WARNING,
            )
        )
        return modules, warnings

    try:
        project = vbe.VBProjects.Item(1)
    except Exception as e:
        warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="VBProjects",
                warning_code="VBA_PROJECT_ERROR",
                message=f"Could not access VBProjects: {e}",
                severity=WarningSeverity.WARNING,
            )
        )
        return modules, warnings

    # Check if project is protected
    project_locked = vba_locked
    if not project_locked:
        try:
            _ = project.VBComponents.Count
        except Exception as e:
            err_str = str(e)
            if any(x in err_str.lower() for x in ("password", "permission", "800a03ec")):
                project_locked = True

    try:
        components = project.VBComponents
        count = components.Count
    except Exception as e:
        if project_locked:
            # Can't enumerate — emit a single blocker warning
            warnings.append(
                ExtractionWarning(
                    object_type="database",
                    object_name="VBAProject",
                    warning_code="VBA_LOCKED",
                    message="VBA project is password-protected; source code not extracted.",
                    severity=WarningSeverity.ERROR,
                )
            )
            return modules, warnings
        warnings.append(
            ExtractionWarning(
                object_type="database",
                object_name="VBComponents",
                warning_code="COM_ERROR",
                message=f"Could not access VBComponents: {e}",
                severity=WarningSeverity.WARNING,
            )
        )
        return modules, warnings

    for i in range(count):
        name = f"<index {i}>"
        try:
            comp = components.Item(i + 1)  # VBComponents is 1-indexed
            try:
                name = str(comp.Name)
            except Exception:
                pass

            module_type = _vba_module_type(comp)
            parent = _parent_object_name(name, module_type)

            source_code: str | None = None
            line_count: int | None = None
            decl_lines: int | None = None

            if not project_locked:
                try:
                    code_mod = comp.CodeModule
                    total = int(code_mod.CountOfLines)
                    line_count = total
                    decl_lines = int(code_mod.CountOfDeclarationLines)
                    if total > 0:
                        source_code = str(code_mod.Lines(1, total))
                except Exception as e:
                    err_str = str(e)
                    if any(x in err_str.lower() for x in ("password", "permission", "800a03ec")):
                        project_locked = True
                    else:
                        warnings.append(
                            ExtractionWarning(
                                object_type="module",
                                object_name=name,
                                warning_code="SOURCE_CODE_ERROR",
                                message=f"Could not read source code: {e}",
                                severity=WarningSeverity.WARNING,
                            )
                        )

            modules.append(
                VBAModule(
                    name=name,
                    module_type=module_type,
                    source_code=source_code,
                    line_count=line_count,
                    declaration_lines=decl_lines,
                    is_vba_locked=project_locked,
                    parent_object=parent,
                )
            )

        except Exception as e:
            warnings.append(
                ExtractionWarning(
                    object_type="module",
                    object_name=name,
                    warning_code="COM_ERROR",
                    message=f"Unexpected error extracting module: {e}",
                    severity=WarningSeverity.ERROR,
                )
            )

    return modules, warnings
