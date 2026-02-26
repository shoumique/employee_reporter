import os
import uuid

from django.conf import settings
from django.http import HttpResponse, JsonResponse
from django.shortcuts import redirect, render
from django.views.decorators.http import require_GET, require_POST

from .utils import (
    REPORT_PRESETS,
    generate_export_excel,
    get_employee_list,
    get_preset_columns,
    load_and_process_excel,
)

# ── Upload ─────────────────────────────────────────────────────────────────

def upload_view(request):
    error = None

    if request.method == "POST":
        uploaded_file = request.FILES.get("excel_file")

        if not uploaded_file:
            error = "কোনো ফাইল নির্বাচন করা হয়নি। একটি Excel ফাইল আপলোড করুন।"
        elif not uploaded_file.name.lower().endswith((".xlsx", ".xls")):
            error = "শুধুমাত্র Excel ফাইল (.xlsx / .xls) আপলোড করুন।"
        else:
            upload_dir = os.path.join(settings.MEDIA_ROOT, "uploads")
            os.makedirs(upload_dir, exist_ok=True)

            ext = os.path.splitext(uploaded_file.name)[1]
            file_id = str(uuid.uuid4())
            file_path = os.path.join(upload_dir, f"{file_id}{ext}")

            with open(file_path, "wb") as fh:
                for chunk in uploaded_file.chunks():
                    fh.write(chunk)

            request.session["excel_path"] = file_path
            request.session["original_filename"] = uploaded_file.name
            return redirect("reporter:configure")

    return render(request, "reporter/upload.html", {"error": error})


# ── Configure ──────────────────────────────────────────────────────────────

def configure_view(request):
    excel_path = request.session.get("excel_path")
    if not excel_path or not os.path.exists(excel_path):
        return redirect("reporter:upload")

    try:
        df, columns = load_and_process_excel(excel_path)
    except Exception as exc:
        return render(
            request,
            "reporter/upload.html",
            {"error": f"ফাইল প্রক্রিয়া করতে ব্যর্থ: {exc}"},
        )

    employees = get_employee_list(df)

    # Build preset → column list mapping for JS
    presets_with_cols = {}
    for key, preset in REPORT_PRESETS.items():
        presets_with_cols[key] = {
            **preset,
            "columns": get_preset_columns(key, columns),
        }

    context = {
        "columns": columns,
        "employees": employees,
        "presets": presets_with_cols,
        "filename": request.session.get("original_filename", "file.xlsx"),
        "row_count": len(df),
    }
    return render(request, "reporter/configure.html", context)


# ── Export ─────────────────────────────────────────────────────────────────

@require_POST
def export_view(request):
    excel_path = request.session.get("excel_path")
    if not excel_path or not os.path.exists(excel_path):
        return redirect("reporter:upload")

    try:
        df, columns = load_and_process_excel(excel_path)
    except Exception:
        return redirect("reporter:upload")

    selected_columns = request.POST.getlist("columns")
    employee_ids = request.POST.getlist("employees")
    report_title = (request.POST.get("report_title") or "Employee Report").strip()

    if not selected_columns:
        selected_columns = columns

    excel_bytes = generate_export_excel(
        df,
        selected_columns,
        employee_ids if employee_ids else None,
        report_title,
    )

    safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in report_title)
    filename = f"{safe_name}.xlsx"
    response = HttpResponse(
        excel_bytes,
        content_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    # Allow the JS fetch client to read the Content-Disposition header
    response["Access-Control-Expose-Headers"] = "Content-Disposition"
    return response


# ── AJAX: preset column list ───────────────────────────────────────────────

@require_GET
def preset_columns_view(request):
    excel_path = request.session.get("excel_path")
    if not excel_path or not os.path.exists(excel_path):
        return JsonResponse({"error": "No file in session"}, status=400)

    preset_key = request.GET.get("preset", "")
    if not preset_key:
        return JsonResponse({"error": "No preset key provided"}, status=400)

    try:
        df, columns = load_and_process_excel(excel_path)
        preset_cols = get_preset_columns(preset_key, columns)
        return JsonResponse({"columns": preset_cols})
    except Exception as exc:
        return JsonResponse({"error": str(exc)}, status=500)
