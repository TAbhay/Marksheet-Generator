from django.urls import path

from web.views import UploadView
from web import views
urlpatterns = [
    path('', UploadView.as_view(), name="index"),
    path('upload', UploadView.as_view(), name="upload"),
    path('individual-marksheet',views.singleMarksheet,name='marksheet'),
    path('concise-marksheet',views.conciseMarksheet),
    path('send-email',views.sendEmail),
    path('*',views.notFound404)
]

