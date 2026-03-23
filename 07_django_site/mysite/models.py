from django.db import models


class UserProfile(models.Model):
    full_name = models.CharField(max_length=255, default="Бартель Андрей Александрович")
    photo = models.ImageField(upload_to="profile_photos/")
    email = models.EmailField(default="aabartel@edu.hse.ru")
    phone = models.CharField(max_length=20, default="+7992*****98")
    resume = models.TextField(default="Ну я")

    def __str__(self):
        return self.full_name


class EducationProgram(models.Model):
    name = models.CharField(max_length=255, default="Экономика")
    link = models.URLField(
        max_length=200,
        blank=True,
        null=True,
        default="https://perm.hse.ru/ba/economics/",
    )
    what_you_will_study = models.TextField(default="Экономику")
    what_you_will_learn = models.TextField(default="Экономить")
    program_benefits = models.TextField(default="Научился экономить")
    future_prospects = models.TextField(default="Много сэкономишь")

    def __str__(self):
        return self.name


class Staff(models.Model):
    full_name = models.CharField(max_length=255, default="Тутынина Ольга Владимировна")
    photo = models.ImageField(
        upload_to="staff_photos/",
        default="https://perm.hse.ru/pubs/share/thumb/905202217:c190x190+0+0:r190x190!",
    )
    email = models.EmailField(default="oshibanova@hse.ru")

    def __str__(self):
        return self.full_name


class Classmate(models.Model):
    full_name = models.CharField(max_length=255, default="Райан Гослинг")
    photo = models.ImageField(
        upload_to="classmate_photos/",
        default="https://img.redbull.com/images/q_auto,f_auto/redbullcom/2016/05/20/1331795954995_2/%D1%80%D0%B0%D0%B9%D0%B0%D0%BD-%D0%B3%D0%BE%D1%81%D0%BB%D0%B8%D0%BD%D0%B3.jpg",
    )
    email = models.EmailField(default="goslya@mail.com")
    phone = models.CharField(max_length=20, default="+7777777777")

    def __str__(self):
        return self.full_name
    

class Page(models.Model):
    title = models.CharField(max_length=100)
    content = models.TextField()
    nav_display = models.BooleanField(default=True)
    nav_position = models.IntegerField(default=0)

    def __str__(self):
        return self.title
