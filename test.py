import pandas as pd
from django.shortcuts import render
from .forms import UploadFileForm

def upload_file(request):
    sheet_data = None
    sheet_names = []
    selected_sheet = None

    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)

        if form.is_valid():
            file = request.FILES['file']

            # Only handle Excel files
            if not file.name.endswith(('.xls', '.xlsx')):
                return render(request, 'upload.html', {
                    'form': form,
                    'error': 'Only Excel files are supported.'
                })

            try:
                # Read all sheets
                excel_data = pd.read_excel(file, sheet_name=None)

                # Extract sheet names
                sheet_names = list(excel_data.keys())

                # Get selected sheet (if any)
                selected_sheet = request.POST.get('sheet_name') or sheet_names[0]

                # Convert selected sheet to HTML
                df = excel_data[selected_sheet]
                sheet_data = df.to_html(classes='table table-bordered', index=False)

            except Exception as e:
                return render(request, 'upload.html', {
                    'form': form,
                    'error': f'Error processing file: {e}'
                })
    else:
        form = UploadFileForm()

    return render(request, 'upload.html', {
        'form': form,
        'sheet_data': sheet_data,
        'sheet_names': sheet_names,
        'selected_sheet': selected_sheet
    })