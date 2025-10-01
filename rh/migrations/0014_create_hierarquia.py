# Generated manually because table was missing
from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('rh', '0013_alter_hierarquia_unique_together'),
    ]

    operations = [
        migrations.CreateModel(
            name='Hierarquia',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('coordenador', models.ForeignKey(
                    on_delete=django.db.models.deletion.CASCADE,
                    related_name='coordenador_set',
                    to=settings.AUTH_USER_MODEL
                )),
                ('supervisor', models.ForeignKey(
                    on_delete=django.db.models.deletion.CASCADE,
                    related_name='supervisor_set',
                    to=settings.AUTH_USER_MODEL
                )),
            ],
            options={
                'verbose_name': 'Hierarquia',
                'verbose_name_plural': 'Hierarquias',
            },
        ),
    ]
