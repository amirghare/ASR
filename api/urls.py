from django.urls import path
from . import views

urlpatterns = [
    path('analyze/', views.analyze_paper, name='analyze'),
    path('generate-pdf/', views.generate_pdf_report, name='generate_pdf'),
    path('generate-word/', views.generate_word_report, name='generate_word'),  # NEW
    path('health/', views.health_check, name='health'),
]