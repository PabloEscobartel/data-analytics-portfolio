from django.contrib import admin
from .models import UserProfile, EducationProgram, Staff, Classmate
from .models import Page

@admin.register(Page)
class PageAdmin(admin.ModelAdmin):
    list_display = ['title', 'nav_display', 'nav_position']
    list_editable = ['nav_display', 'nav_position']
    search_fields = ['title', 'content']
    list_filter = ['nav_display']
admin.site.register(UserProfile)
admin.site.register(EducationProgram)
admin.site.register(Staff)
admin.site.register(Classmate)
