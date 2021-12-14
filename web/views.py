from django.shortcuts import render
from django.views import generic
from django.contrib import messages
from web.forms import UploadForm
from django.shortcuts import render, redirect
from django.urls import reverse_lazy
import time
from django.http import JsonResponse
from .models import Photo
from django.http import HttpResponse
from web import controllers
from django.conf import settings
from django.core.mail import send_mail
from pathlib import Path
import os.path
import os
import smtplib
from email.message import EmailMessage
import glob


def handleAjaxUpload(request):
    try:
        filePath = ""
        form = UploadForm(request.POST, request.FILES)
        if form.is_valid():
            print("hwew",form)
            form.save()
            return JsonResponse({
                "code": 201,
                "message": "Image uploaded successfully"
            })
        else:
            return JsonResponse({
                "status": 400,
                "message": form.errors["image"]
            })
    except Exception as e:
        return JsonResponse({
            "code": 500,
            "message": e
        })        


class UploadView(generic.View):
    def get(self, *args, **kwargs):
        form = UploadForm()
        imageSrc = "/media/default.png"
        photo = Photo.objects.last()
        if photo:
            imageSrc = photo.image.url
        return render(self.request, "web/index.html",  {
            "form": form,
            "imageSrc": imageSrc
        })

    def post(self, *args, **kwargs):
        if self.request.is_ajax():
            return handleAjaxUpload(self.request)
        form = UploadForm(self.request.POST, self.request.FILES)
        if form.is_valid():
            filename = self.request.FILES['image'].name
            filePath ="media/input/"+filename
            print(filePath)
            if os.path.exists(filePath):
                os.remove(filePath)
            form.save()
            messages.success(self.request, "Image Uploaded")
        else:
            messages.error(self.request, "Failed uploading image")
        
        return redirect(reverse_lazy("index"))


def singleMarksheet(request):
    try:
        positive = request.GET.get('positive')
        negative = request.GET.get('negative')
        controllers.makeOutputDir()
        controllers.generate_result(positive,negative)
        controllers.individual_marksheet()
        obj = {"a": 1, "b": 2}
        return render(request, "web/success.html",{"success":"Success - Individual Marksheet"})
    except Exception as e:
        return render(request, "web/error.html",{"success":e}) 
    
def conciseMarksheet(request):
    try:
        print(request.GET.get('positive'))
        positive = request.GET.get('positive')
        negative = request.GET.get('negative')
        controllers.makeOutputDir()
        controllers.generate_result(positive,negative)
        controllers.concise_marksheet()
        return render(request, "web/success.html",{"success":"Success - Concise Marksheet"})
    except Exception as e:
        return render(request, "web/error.html",{"success":e}) 

# server email port password
serverSMtp = "server"
port = "port"
email  = "email"
password ="password" 


def sendEmail(request):
    try:
        controllers.sendemail()
        return render(request, "web/success.html",{"success":"Success - Mail Sent to student"})
    except Exception as e:
        return JsonResponse({
            "code": 500,
            "message": e
        }) 
def notFound404(request):
    try:
        
        return render(request, "web/notFound.html",{"success":"Error : 404"})
    except Exception as e:
        return render(request, "web/error.html",{"success":e}) 

