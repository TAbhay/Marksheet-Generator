from django.db import models
from django.conf import settings
from django.core.exceptions import ValidationError
# Create your models here.

MAX_SIZE_ALLOWED_IN_KB = 500;

def validate_size(FileField):
    size = FileField.file.size
    if size > 1024 * MAX_SIZE_ALLOWED_IN_KB:
        raise ValidationError(
            message= "File size should not exceed %(value)s KB",
            params= {
                "value": MAX_SIZE_ALLOWED_IN_KB
            },
            code= 123
        ) 
    if not FileField.file.name.endswith('.csv'):
        raise ValidationError(
            message= "File size should not exceed %(value)s KB",
            params= {
                "value": MAX_SIZE_ALLOWED_IN_KB
            },
            code= 123
        )    

class Photo(models.Model):
    image = models.FileField(upload_to = "input", validators= [validate_size])
