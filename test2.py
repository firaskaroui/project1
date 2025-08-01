import pandas as pd
import base64
from io import BytesIO
from django.shortcuts import render
from .forms import UploadFileForm

def upload_file(request):
    sheet_data = None
    sheet_names = []
    selected_sheet = None
    form = UploadFileForm()

    if request.method == 'POST':
        # Sheet dropdown submission
        if 'sheet_name' in request.POST and 'excel_file' in request.session:
            # Load file from session
            file_content = base64.b64decode(request.session['excel_file'])
            excel_io = BytesIO(file_content)

            try:
                excel_data = pd.read_excel(excel_io, sheet_name=None)
                sheet_names = list(excel_data.keys())
                selected_sheet = request.POST['sheet_name']
                df = excel_data[selected_sheet]
                sheet_data = df.to_html(classes='table table-striped table-bordered', index=False)
            except Exception as e:
                return render(request, 'upload.html', {
                    'form': form,
                    'error': f'Error reading sheet: {e}'
                })

        # New file upload
        elif 'file' in request.FILES:
            form = UploadFileForm(request.POST, request.FILES)
            if form.is_valid():
                file = request.FILES['file']

                if not file.name.endswith(('.xls', '.xlsx')):
                    return render(request, 'upload.html', {
                        'form': form,
                        'error': 'Only Excel files are supported.'
                    })

                # Save file content in session (as base64)
                content = file.read()
                request.session['excel_file'] = base64.b64encode(content).decode('utf-8')

                excel_io = BytesIO(content)
                excel_data = pd.read_excel(excel_io, sheet_name=None)
                sheet_names = list(excel_data.keys())
                selected_sheet = sheet_names[0]
                df = excel_data[selected_sheet]
                sheet_data = df.to_html(classes='table table-striped table-bordered', index=False)

    return render(request, 'upload.html', {
        'form': form,
        'sheet_data': sheet_data,
        'sheet_names': sheet_names,
        'selected_sheet': selected_sheet
    })