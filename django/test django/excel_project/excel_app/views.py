from django.shortcuts import render, redirect
from .forms import UploadFilesForm
from django.http import HttpResponse
import pandas as pd
import json
import os
import creation_fichor

def handle_uploaded_files(file1, file2):
    file_path1 = file1.name
    file_path2 = file2.name
    with open(file_path1, 'wb+') as destination:
        for chunk in file1.chunks():
            destination.write(chunk)
    with open(file_path2, 'wb+') as destination:
        for chunk in file2.chunks():
            destination.write(chunk)
    df1 = pd.read_excel(file_path1)
    df2 = pd.read_excel(file_path2)
    return df1, df2, file_path1, file_path2

def upload_files(request):
    if request.method == 'POST':
        form = UploadFilesForm(request.POST, request.FILES)
        if form.is_valid():
            df1, df2, file_path1, file_path2 = handle_uploaded_files(request.FILES['file1'], request.FILES['file2'])
            df1_json = df1.to_json(orient='split')
            df2_json = df2.to_json(orient='split')
            # Save original file names to the session
            request.session['file_path1'] = file_path1
            request.session['file_path2'] = file_path2
            return render(request, 'display.html', {
                'df1_json': df1_json,
                'df2_json': df2_json,
                'file_path1': file_path1,
                'file_path2': file_path2
            })
    else:
        form = UploadFilesForm()
    return render(request, 'upload.html', {'form': form})

def modify_data(request):
    if request.method == 'POST':
        df1_json = request.POST.get('df1_json')
        df2_json = request.POST.get('df2_json')
        df1 = pd.read_json(df1_json, orient='split')
        df2 = pd.read_json(df2_json, orient='split')
        # Convert DataFrames to HTML and include editing capabilities
        df1_html = df1.to_html(classes='table table-striped', table_id='editableTable1')
        df2_html = df2.to_html(classes='table table-striped', table_id='editableTable2')
        return render(request, 'modify.html', {
            'df1_html': df1_html,
            'df2_html': df2_html,
            'df1_json': df1_json,
            'df2_json': df2_json,
            'file_path1': request.POST.get('file_path1'),
            'file_path2': request.POST.get('file_path2')
        })
    return redirect('upload_files')

def save_changes(request):
    if request.method == 'POST':
        df1_json = request.POST.get('df1_json')
        df2_json = request.POST.get('df2_json')
        df1 = pd.read_json(df1_json, orient='split')
        df2 = pd.read_json(df2_json, orient='split')
        updated_data1 = json.loads(request.POST.get('updated_data1'))
        updated_data2 = json.loads(request.POST.get('updated_data2'))
        df1 = pd.DataFrame(updated_data1)
        df2 = pd.DataFrame(updated_data2)
        file_path1 = request.POST.get('file_path1')
        file_path2 = request.POST.get('file_path2')
        df1.to_excel(file_path1, index=False)
        df2.to_excel(file_path2, index=False)
        correct_exel(request)
        return render(request, 'success.html', {
            'file_path1': file_path1,
            'file_path2': file_path2
        })
    return redirect('upload_files')

def download_files(request):
    file_path1 = request.GET.get('file_path1')
    file_path2 = request.GET.get('file_path2')

    if file_path1:
        with open(file_path1, 'rb') as file1:
            response = HttpResponse(file1.read(), content_type='application/vnd.ms-excel')
            response['Content-Disposition'] = f'attachment; filename={os.path.basename(file_path1)}'
            return response

    if file_path2:
        with open(file_path2, 'rb') as file2:
            response = HttpResponse(file2.read(), content_type='application/vnd.ms-excel')
            response['Content-Disposition'] = f'attachment; filename={os.path.basename(file_path2)}'
            return response

    return redirect('upload_files')

def correct_exel(request):
    # Retrieve file paths from session
    file_path1 = request.session.get('file_path1')
    file_path2 = request.session.get('file_path2')
    
    if not file_path1 or not file_path2:
        return redirect('upload_files')
    
    # Read the Excel files
    df1 = pd.read_excel(file_path1)
    df2 = pd.read_excel(file_path2)
    
    # Check and remove columns after the pattern [0, 1, 2] in the first row
    def remove_pattern_columns(df):
        if len(df) > 0:
            first_row = df.iloc[0, :].tolist()
            pattern = [0, 1, 2]
            for i in range(len(first_row) - len(pattern) + 1):
                if first_row[i:i+len(pattern)] == pattern:
                    # Remove all columns after the pattern
                    df = df.iloc[:, :i+len(pattern)]
                    break
        return df
    
    df1 = remove_pattern_columns(df1)
    df2 = remove_pattern_columns(df2)
    
    # Write the modified DataFrames back to the original files
    df1.to_excel(file_path1, header=False, index=False)
    df2.to_excel(file_path2, header=False, index=False)

def creation_fichor_html():
try:
    file1=creation_fichor.lire_fichier_entree(file_path1, 'Feuil1')
    file2=creation_fichor.lire_fichier_entree(file_path2, 'Feuil1')
    creation_fichor.ecrire_fichier_sortie(file1, file2, 'fichier_horaire_enseignant_sortie.xlsx', 'fichier_referenciel.xlsx')
