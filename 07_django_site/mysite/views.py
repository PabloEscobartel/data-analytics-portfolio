from django.shortcuts import render, get_object_or_404
from .models import UserProfile, EducationProgram, Staff, Classmate, Page


def index(request):
    user_profile = UserProfile.objects.first()
    pages = Page.objects.filter(nav_display=True).order_by('nav_position')
    return render(request, 'mysite/index.html', {'profile': user_profile, 'pages': pages})

def education_program(request):
    education_program = EducationProgram.objects.first()
    pages = Page.objects.filter(nav_display=True).order_by('nav_position')
    return render(request, 'mysite/education_program.html', {'program': education_program, 'pages': pages})

def staff(request):
    staff_members = Staff.objects.all()
    pages = Page.objects.filter(nav_display=True).order_by('nav_position')
    return render(request, 'mysite/staff.html', {'staff': staff_members, 'pages': pages})

def classmates(request):
    classmates = Classmate.objects.all()
    pages = Page.objects.filter(nav_display=True).order_by('nav_position')
    return render(request, 'mysite/classmates.html', {'classmates': classmates, 'pages': pages})


def page_view(request, page_id):
    page = get_object_or_404(Page, id=page_id)
    pages = Page.objects.filter(nav_display=True).order_by('nav_position')
    return render(request, 'mysite/page.html', {'page': page, 'pages': pages})
