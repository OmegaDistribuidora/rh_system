# Fixed migration (noop)
from django.db import migrations

class Migration(migrations.Migration):

    dependencies = [
        ('rh', '0012_alter_hierarquia_options_and_more'),
    ]

    operations = [
        # Nenhuma operação necessária, já que o unique_together nunca existiu
    ]
