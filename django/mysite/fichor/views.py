from django.shortcuts import render
from django.shortcuts import render


def home(request):
    return render(request, 'home.html')

def fichor(request):
    return render(request, 'fichor.html')

def display_fichor(request):
    
    if request.method == 'POST':
        # Handle file uploads
        Code_Etape = request.FILES['file1']
        List_Ens = request.FILES['file2']
        
  
    return render(request, 'display_fichor.html', {'Code_Etape': Code_Etape, 'List_Ens': List_Ens})

def download_fich(request):
    return render(request, 'download_fich.html')

def list_ens(request):
    return render(request, 'list_ens.html')

def display_list(request):
    if request.method == 'POST':
        # Handle file uploads
        referenciel = request.FILES['file1']
        hor_enseignant = request.FILES['file2']
        

    return render(request, 'display_list.html', {'referenciel': referenciel, 'hor_enseignant': hor_enseignant})