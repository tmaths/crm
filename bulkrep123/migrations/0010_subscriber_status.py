from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('bulkrep', '0009_auto_20251111_1150'),
    ]

    operations = [
        migrations.AddField(
            model_name='subscriber',
            name='status',
            field=models.CharField(max_length=20, choices=[('active','Active'),('inactive','Inactive'),('pending','Pending')], default='pending'),
        ),
    ]
