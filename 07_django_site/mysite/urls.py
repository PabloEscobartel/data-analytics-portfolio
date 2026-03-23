from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('education_program/', views.education_program, name='education_program'),
    path('staff/', views.staff, name='staff'),
    path('classmates/', views.classmates, name='classmates'),
    path('page/<int:page_id>/', views.page_view, name='page'),
]