from django.db import models

# Create your models here.


class YearData(models.Model):
    year = models.IntegerField()
    is_leap = models.BooleanField()
    century = models.CharField(max_length=20)
    era = models.CharField(max_length=10, default='н. э.')
