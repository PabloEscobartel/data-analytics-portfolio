from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='indexleap'),
    path('results/', views.results, name='results'),
    path('all_data/', views.all_data, name='all_data'),
]
