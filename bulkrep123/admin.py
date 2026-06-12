from django.contrib import admin
from .models import Usagereport, ReportGeneration, SubscriberProductRate, ENQUIRY_RATES, Subscriber, KeySubscriber
from django.utils.html import format_html
from django.urls import reverse
from django.utils.safestring import mark_safe
from django.utils import timezone
from django import forms
from decimal import Decimal


class UsagereportAdmin(admin.ModelAdmin):
    list_display = ('SubscriberName', 'DetailsViewedDate', 'ProductName', 'SystemUser','SearchIdentity')
    list_filter = ('DetailsViewedDate', 'ProductName', 'SearchIdentity')
    search_fields = ('SubscriberName', 'ProductName', 'SystemUser', 'SearchIdentity')
    date_hierarchy = 'DetailsViewedDate'
    ordering = ('-DetailsViewedDate', 'SubscriberName')
    list_per_page = 50


class ReportGenerationAdmin(admin.ModelAdmin):
    list_display = ('user', 'report_type', 'status', 'formatted_generated_at', 'formatted_completed_at', 'subscriber_name', 'duration_display')
    list_filter = ('report_type', 'status', 'generated_at')
    search_fields = ('user__username', 'subscriber_name', 'generator')
    date_hierarchy = 'generated_at'
    readonly_fields = ('generated_at', 'completed_at', 'duration_display')
    list_per_page = 50
    
    def formatted_generated_at(self, obj):
        if obj.generated_at:
            local_time = timezone.localtime(obj.generated_at)
            return local_time.strftime('%Y-%m-%d %H:%M:%S')
        return "-"
    formatted_generated_at.short_description = 'Generated At (Local)'
    formatted_generated_at.admin_order_field = 'generated_at'
    
    def formatted_completed_at(self, obj):
        if obj.completed_at:
            local_time = timezone.localtime(obj.completed_at)
            return local_time.strftime('%Y-%m-%d %H:%M:%S')
        return "-"
    formatted_completed_at.short_description = 'Completed At (Local)'
    formatted_completed_at.admin_order_field = 'completed_at'
    
    def duration_display(self, obj):
        if obj.duration is not None:
            return f"{obj.duration:.2f} seconds"
        return "-"
    duration_display.short_description = 'Duration'


class SubscriberProductRateForm(forms.ModelForm):
    class Meta:
        model = SubscriberProductRate
        fields = '__all__'
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        # Get unique subscriber names from Usagereport
        subscriber_choices = [(name, name) for name in 
                            Usagereport.objects.values_list('SubscriberName', flat=True)
                            .distinct().order_by('SubscriberName')]
        
        # Get unique product names from Usagereport
        product_choices = [(name, name) for name in 
                          Usagereport.objects.values_list('ProductName', flat=True)
                          .distinct().order_by('ProductName')]
        
        # Create choices for predefined rates
        rate_choices = [(str(rate), f"{product.replace('_', ' ').title()} - ₦{rate:,.2f}") 
                       for product, rate in ENQUIRY_RATES.items()]
        rate_choices.append(('custom', 'Custom Rate'))
        
        # Set up subscriber dropdown
        self.fields['subscriber_name'] = forms.ChoiceField(
            choices=[('', 'Select Subscriber')] + subscriber_choices,
            widget=forms.Select(attrs={'class': 'form-control'})
        )
        
        # Set up product dropdown
        self.fields['product_name'] = forms.ChoiceField(
            choices=[('', 'Select Product')] + product_choices,
            widget=forms.Select(attrs={'class': 'form-control'})
        )
        
        # Set up rate dropdown with custom option
        self.fields['rate_choice'] = forms.ChoiceField(
            choices=[('', 'Select Rate')] + rate_choices,
            required=False,
            label='Predefined Rates',
            widget=forms.Select(attrs={
                'class': 'form-control',
                'onchange': 'toggleCustomRate(this)'
            })
        )
        
        # Keep the original rate field for custom input
        self.fields['rate'].widget.attrs.update({
            'class': 'form-control',
            'placeholder': 'Enter custom rate'
        })
        
    def clean(self):
        cleaned_data = super().clean()
        rate_choice = cleaned_data.get('rate_choice')
        rate = cleaned_data.get('rate')
        
        if rate_choice and rate_choice != 'custom':
            # Use predefined rate
            cleaned_data['rate'] = Decimal(rate_choice)
        elif rate_choice == 'custom':
            # For custom rate, ensure rate field has a value
            if not rate:
                raise forms.ValidationError('Please enter a custom rate.')
        elif not rate_choice:
            # If no rate_choice is selected, rate field must have a value
            if not rate:
                raise forms.ValidationError('Please select a predefined rate or enter a custom rate.')
            
        return cleaned_data


class SubscriberProductRateAdmin(admin.ModelAdmin):
    form = SubscriberProductRateForm
    list_display = ('subscriber_name', 'product_name', 'rate', 'formatted_rate')
    list_filter = ('subscriber_name', 'product_name',)
    search_fields = ('subscriber_name', 'product_name')
    list_editable = ('rate',)
    ordering = ('subscriber_name', 'product_name')
    
    class Media:
        js = ('admin/js/subscriber_rate_admin.js',)
        css = {
            'all': ('admin/css/subscriber_rate_admin.css',)
        }
    
    def formatted_rate(self, obj):
        try:
            # Convert to Decimal if it's a string
            rate_value = Decimal(str(obj.rate)) if obj.rate else Decimal('0')
            return f"₦{rate_value:,.2f}"
        except (ValueError, TypeError):
            return f"₦{obj.rate}"
    formatted_rate.short_description = 'Rate (₦)'
    formatted_rate.admin_order_field = 'rate'

    def response_change(self, request, obj):
        # Custom message for single object change
        self.message_user(
            request,
            f"{obj.subscriber_name}, {obj.product_name} rate changed to ₦{obj.rate:,.2f}",
            level='success'
        )
        return super().response_change(request, obj)

    def response_action(self, request, queryset):
        # Custom message for bulk actions
        changed = queryset.count()
        if changed == 1:
            obj = queryset.first()
            self.message_user(
                request,
                f"{obj.subscriber_name}, {obj.product_name} rate changed to ₦{obj.rate:,.2f}",
                level='success'
            )
        elif changed > 1:
            names = ", ".join(f"{obj.subscriber_name}, {obj.product_name}" for obj in queryset)
            self.message_user(
                request,
                f"Rates changed for: {names}",
                level='success'
            )
        return super().response_action(request, queryset)


class SubscriberAdmin(admin.ModelAdmin):
    list_display = ('name', 'status', 'managed_by', 'contact_person', 'email', 'phone_number')
    list_filter = ('status', 'managed_by')
    search_fields = ('name', 'contact_person', 'email', 'phone_number')
    ordering = ('name',)


class KeySubscriberAdminForm(forms.ModelForm):
    """Custom form for KeySubscriber with dropdown populated from database."""
    class Meta:
        model = KeySubscriber
        fields = ['subscriber_name']
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Get unique subscriber names from Usagereport
        subscriber_choices = [(name, name) for name in 
                            Usagereport.objects.values_list('SubscriberName', flat=True)
                            .distinct().order_by('SubscriberName')]
        
        # Set up subscriber dropdown with search
        self.fields['subscriber_name'] = forms.ChoiceField(
            choices=[('', 'Select Subscriber')] + subscriber_choices,
            widget=forms.Select(attrs={
                'class': 'form-control',
                'style': 'width: 400px;'
            }),
            label='Subscriber Name'
        )


class KeySubscriberAdmin(admin.ModelAdmin):
    """Admin interface for managing key/priority subscribers."""
    form = KeySubscriberAdminForm
    list_display = ('subscriber_name', 'added_at', 'added_by')
    search_fields = ('subscriber_name',)
    ordering = ('subscriber_name',)
    readonly_fields = ('added_at',)
    
    def save_model(self, request, obj, form, change):
        if not change:  # Only set added_by when creating new entry
            obj.added_by = request.user
        super().save_model(request, obj, form, change)


# Register your models with the custom admin classes
admin.site.register(Usagereport, UsagereportAdmin)
admin.site.register(ReportGeneration, ReportGenerationAdmin)
admin.site.register(SubscriberProductRate, SubscriberProductRateAdmin)
admin.site.register(Subscriber, SubscriberAdmin)
admin.site.register(KeySubscriber, KeySubscriberAdmin)
