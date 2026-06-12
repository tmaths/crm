from django.core.management.base import BaseCommand
from django.contrib.auth.models import Group, Permission
from django.contrib.contenttypes.models import ContentType
from bulkrep.models import Subscriber

class Command(BaseCommand):
    help = 'Creates the initial Manager group and assigns permissions.'

    def handle(self, *args, **options):
        # Create the Manager group
        manager_group, created = Group.objects.get_or_create(name='Manager')
        if created:
            self.stdout.write(self.style.SUCCESS('Successfully created the "Manager" group.'))
        else:
            self.stdout.write('The "Manager" group already exists.')

        # Get the content type for the Subscriber model
        try:
            subscriber_content_type = ContentType.objects.get_for_model(Subscriber)
        except ContentType.DoesNotExist:
            self.stderr.write(self.style.ERROR('The Subscriber model does not exist. Please create it first.'))
            return

        # Define the permissions for the Manager group
        permissions = [
            Permission.objects.get(codename='view_subscriber', content_type=subscriber_content_type),
            Permission.objects.get(codename='change_subscriber', content_type=subscriber_content_type),
        ]

        # Assign the permissions to the Manager group
        manager_group.permissions.set(permissions)
        self.stdout.write(self.style.SUCCESS('Successfully assigned permissions to the "Manager" group.'))