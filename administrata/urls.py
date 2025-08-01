from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='administrata_home'),
    path('01/', views.albi_fashion, name='albi_fashion'),
    path('01/daily-z-report/', views.daily_z_report_albi_fashion, name='daily_z_report_albi_fashion'),
    path('01/monthly-z-report/', views.monthly_z_report_albi_fashion, name='monthly_z_report_albi_fashion'),
    path('01/edit-z-report/', views.edit_z_report_albi_fashion, name='edit_z_report_albi_fashion'),
    path('01/procedura/', views.procedura_albi_fashion, name='procedura_albi_fashion'),

    path('02/', views.ardi_co, name='ardi_co'),
    path('02/daily-z-report/', views.daily_z_report_ardi_co, name='daily_z_report_ardi_co'),
    path('02/monthly-z-report/', views.monthly_z_report_ardi_co, name='monthly_z_report_ardi_co'),
    path('02/edit-z-report/', views.edit_z_report_ardi_co, name='edit_z_report_ardi_co'),
    path('02/procedura/', views.procedura_ardi_co, name='procedura_ardi_co'),

    path('03/', views.nan, name='nan'),
    path('03/daily-z-report/', views.daily_z_report_nan, name='daily_z_report_nan'),
    path('03/monthly-z-report/', views.monthly_z_report_nan, name='monthly_z_report_nan'),
    path('03/edit-z-report/', views.edit_z_report_nan, name='edit_z_report_nan'),
    path('03/procedura/', views.procedura_nan, name='procedura_nan'),
    
    path('04/', views.obe, name='obe'),
    path('04/daily-z-report/', views.daily_z_report_obe, name='daily_z_report_obe'),
    path('04/monthly-z-report/', views.monthly_z_report_obe, name='monthly_z_report_obe'),
    path('04/edit-z-report/', views.edit_z_report_obe, name='edit_z_report_obe'),
    path('04/procedura/', views.procedura_obe, name='procedura_obe'),
    
    path('05/', views.dndo, name='dndo'),
    path('05/daily-z-report/', views.daily_z_report_dndo, name='daily_z_report_dndo'),
    path('05/monthly-z-report/', views.monthly_z_report_dndo, name='monthly_z_report_dndo'),
    path('05/edit-z-report/', views.edit_z_report_dndo, name='edit_z_report_dndo'),
    path('05/procedura/', views.procedura_dndo, name='procedura_dndo'),

    path('06/', views.jaroma, name='jaroma'),
    path('06/daily-z-report/', views.daily_z_report_jaroma, name='daily_z_report_jaroma'),
    path('06/monthly-z-report/', views.monthly_z_report_jaroma, name='monthly_z_report_jaroma'),
    path('06/edit-z-report/', views.edit_z_report_jaroma, name='edit_z_report_jaroma'),
    path('06/procedura/', views.procedura_jaroma, name='procedura_jaroma'),

    path('07/', views.albi_fashion_retail, name='albi_fashion_retail'),
    path('07/daily-z-report/', views.daily_z_report_albi_fashion_retail, name='daily_z_report_albi_fashion_retail'),
    path('07/monthly-z-report/', views.monthly_z_report_albi_fashion_retail, name='monthly_z_report_albi_fashion_retail'),
    path('07/edit-z-report/', views.edit_z_report_albi_fashion_retail, name='edit_z_report_albi_fashion_retail'),
    path('07/procedura/', views.procedura_albi_fashion_retail, name='procedura_albi_fashion_retail'),

    path('08/', views.ran, name='ran'),
    path('08/daily-z-report/', views.daily_z_report_ran, name='daily_z_report_ran'),
    path('08/monthly-z-report/', views.monthly_z_report_ran, name='monthly_z_report_ran'),
    path('08/edit-z-report/', views.edit_z_report_ran, name='edit_z_report_ran'),
    path('08/procedura/', views.procedura_ran, name='procedura_ran'),
]
