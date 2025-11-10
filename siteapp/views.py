from django.shortcuts import render

def home(request):
    return render(request, 'siteapp/home.html', {"title": "AppMaps Web"})
