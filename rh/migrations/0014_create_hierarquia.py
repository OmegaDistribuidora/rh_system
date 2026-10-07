# Compatibility migration.
#
# Hierarquia is already created by migration 0011. Keeping a second
# CreateModel here makes new databases fail with "table already exists".
from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ("rh", "0013_alter_hierarquia_unique_together"),
    ]

    operations = [
        migrations.AlterUniqueTogether(
            name="hierarquia",
            unique_together=set(),
        ),
    ]
