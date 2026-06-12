from django.urls import path
from . import views


app_name = 'bulkrep'
 
urlpatterns = [
    # path('',views.home, name='home'),
    path('', views.dashboard, name='home'),
    path('single-report/', views.single_report, name='single_report'),
    path('bulk-report/', views.bulk_report, name='bulk_report'),
    path('dashboard/', views.dashboard, name='dashboard'),
    path('dashboard-api/', views.dashboard_api, name='dashboard_api'),
    path('subscriber-selection/', views.subscriber_selection, name='subscriber_selection'),
    path('my-subscribers/', views.my_subscribers, name='my_subscribers'),
    path('my-subscribers/new/', views.subscriber_create, name='subscriber_create'),
    path('my-subscribers/<int:pk>/edit/', views.subscriber_update, name='subscriber_update'),
    path('reassign-subscribers/', views.admin_reassign_subscribers, name='admin_reassign_subscribers'),
    path('manage-subscribers/', views.admin_manage_subscribers, name='admin_manage_subscribers'),
    path('download-subscribers-pdf/', views.download_subscribers_pdf, name='download_subscribers_pdf'),
    path('download-churned-subscribers/', views.download_churned_subscribers, name='download_churned_subscribers'),
    path('download-new-subscribers/', views.download_new_subscribers, name='download_new_subscribers'),
    path('download-top-subscribers-csv/', views.download_top_subscribers_csv, name='download_top_subscribers_csv'),
    path('download-top-subscribers-summary/', views.download_top_subscribers_summary, name='download_top_subscribers_summary'),
    path('download-bottom-subscribers-csv/', views.download_bottom_subscribers_csv, name='download_bottom_subscribers_csv'),
    path('api/new-subscribers-trend/', views.new_subscribers_trend_api, name='new_subscribers_trend_api'),
    path('api/usage-trends/', views.usage_trends_api, name='usage_trends_api'),
    path('download-usage-trends-excel/', views.download_usage_trends_excel, name='download_usage_trends_excel'),
    path('download-active-subscribers/', views.download_active_subscribers, name='download_active_subscribers'),
    path('subscriber-performance/', views.subscriber_performance, name='subscriber_performance'),
    path('api/subscriber-performance/', views.subscriber_performance_api, name='subscriber_performance_api'),
    path('submission-tracking/', views.submission_tracking, name='submission_tracking'),
    path('api/submission-tracking/', views.submission_tracking_api, name='submission_tracking_api'),
    path('download-non-submitters/', views.download_non_submitters, name='download_non_submitters'),
    path('download-missing-from-submitteddata/', views.download_missing_from_submitteddata, name='download_missing_from_submitteddata'),
]

