from django.db import models

# Create your models here.

class ReceiptImage(models.Model):
    image = models.ImageField(upload_to='receipts/')
    uploaded_at = models.DateTimeField(auto_now_add=True)
    differences_checked = models.BooleanField(default=False)

    def __str__(self):
        return f"Receipt uploaded at {self.uploaded_at.strftime('%Y-%m-%d %H:%M:%S')}"
