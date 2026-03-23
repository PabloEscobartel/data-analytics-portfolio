from django.shortcuts import render, redirect
from django.db.models import Avg, Min, Max, Count, StdDev, Sum
from .models import YearData
from .forms import YearForm

def is_leap_year(year):
    if year > 0:
        if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0):
            return True
    else:
        if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0):
            return True
    return False

def get_century_and_era(year):
    if year > 0:
        century = (year - 1) // 100 + 1
        era = 'AD'
    else:
        century = (-year) // 100 + 1
        era = 'BC'
    return f"{century} век", era

def index(request):
    if request.method == 'POST':
        form = YearForm(request.POST)
        if form.is_valid():
            year = form.cleaned_data['year']
            is_leap = is_leap_year(year)
            century, era = get_century_and_era(year)
            YearData.objects.create(year=year, is_leap=is_leap, century=century, era=era)
            return redirect('results')
    else:
        form = YearForm()
    return render(request, 'leap_year/indexleap.html', {'form': form})

def results(request):
    latest_entry = YearData.objects.latest('id')
    return render(request, 'leap_year/results.html', {'entry': latest_entry})

def all_data(request):
    sort_by = 'id'
    all_entries = YearData.objects.all().order_by(sort_by)

    statistics = {
        'count': all_entries.aggregate(Count('year'))['year__count'],
        'avg': all_entries.aggregate(Avg('year'))['year__avg'],
        'min': all_entries.aggregate(Min('year'))['year__min'],
        'max': all_entries.aggregate(Max('year'))['year__max'],
        'stddev': all_entries.aggregate(StdDev('year'))['year__stddev'],
        'sum': all_entries.aggregate(Sum('year'))['year__sum'],
    }

    return render(request, 'leap_year/all_data.html', {'all_entries': all_entries, 'statistics': statistics})
