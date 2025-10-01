from django.db import migrations, models

class Migration(migrations.Migration):

    dependencies = [
        ("rh", "0015_admissao_status_desligamento_status_distrato_status"),
    ]

    operations = [
        migrations.AlterField(
            model_name="desligamento",
            name="codigo",
            field=models.CharField("Código", max_length=20, db_index=True),
        ),
        migrations.AlterField(
            model_name="admissao",
            name="codigo",
            field=models.CharField("Código RCA", max_length=20, db_index=True),
        ),
    ]
