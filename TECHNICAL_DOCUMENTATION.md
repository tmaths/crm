# Bulk Report Dashboard - Technical Documentation

## Table of Contents
1. [Architecture Overview](#architecture-overview)
2. [Technology Stack](#technology-stack)
3. [Development Environment Setup](#development-environment-setup)
4. [Project Structure](#project-structure)
5. [Database Schema](#database-schema)
6. [API Documentation](#api-documentation)
7. [Frontend Architecture](#frontend-architecture)
8. [Backend Implementation](#backend-implementation)
9. [Caching Strategy](#caching-strategy)
10. [Error Handling & Logging](#error-handling--logging)
11. [Performance Optimization](#performance-optimization)
12. [Deployment Guide](#deployment-guide)
13. [Testing](#testing)
14. [Security Considerations](#security-considerations)
15. [Troubleshooting](#troubleshooting)

---

## Architecture Overview

### System Design

The Bulk Report Dashboard follows a **3-tier architecture**:

```
┌─────────────────────────────────────────────────────┐
│              Frontend Layer (Browser)                │
│         HTML5, CSS3, JavaScript (Chart.js)           │
└──────────────────┬──────────────────────────────────┘
                   │ HTTP/AJAX
┌──────────────────▼──────────────────────────────────┐
│         Application Layer (Django)                   │
│   Views, Templates, URL Routing, Business Logic      │
├─────────────────────────────────────────────────────┤
│         Caching Layer (Database Cache)               │
└──────────────────┬──────────────────────────────────┘
                   │ ODBC
┌──────────────────▼──────────────────────────────────┐
│     Data Layer (Microsoft SQL Server)                │
│   Models, ORM, Database Connection Pool              │
└─────────────────────────────────────────────────────┘
```

### Key Components

| Component | Purpose | Technology |
|-----------|---------|-----------|
| **Frontend** | User interface and interactivity | HTML5, CSS3, JavaScript, Chart.js |
| **Backend** | Business logic and API endpoints | Django 2.1.15 |
| **Database** | Data persistence | Microsoft SQL Server |
| **Cache** | Performance optimization | Django Database Cache |
| **Task Queue** | Async processing | Celery |
| **File Processing** | Excel report generation | OpenPyXL |

---

## Technology Stack

### Backend
- **Framework**: Django 2.1.15
- **ORM**: Django ORM with MSSQL backend
- **Task Queue**: Celery for background tasks
- **Database Driver**: pyodbc with ODBC Driver 17 for SQL Server

### Frontend
- **Markup**: HTML5
- **Styling**: CSS3 (Responsive Design)
- **Scripting**: Vanilla JavaScript
- **Charting**: Chart.js for interactive visualizations
- **Data Format**: JSON

### Database
- **Engine**: Microsoft SQL Server
- **Connection**: ODBC Driver 17 for SQL Server
- **Connection Method**: Trusted Connection (Windows Authentication)

### Development Tools
- **Package Manager**: pip
- **Version Control**: Git
- **Virtual Environment**: venv
- **Excel Processing**: OpenPyXL

---

## Development Environment Setup

### System Requirements

| Requirement | Version | Notes |
|------------|---------|-------|
| Python | 3.8+ | Required for Django and dependencies |
| SQL Server | 2016+ | For data persistence |
| ODBC Driver | 17 for SQL Server | For database connectivity |
| Git | Latest | For version control |

### Step-by-Step Setup

#### 1. Prerequisites Installation

**On Windows:**
```bash
# Download and install ODBC Driver 17 for SQL Server
# https://docs.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server

# Verify installation
odbcad32.exe  # Open ODBC Data Source Administrator
```

**On Linux (Ubuntu/Debian):**
```bash
curl https://packages.microsoft.com/keys/microsoft.asc | apt-key add -
curl https://packages.microsoft.com/config/ubuntu/20.04/prod.list > /etc/apt/sources.list.d/mssql-release.list
apt-get update
apt-get install -y msodbcsql17
```

#### 2. Clone and Setup Repository

```bash
# Clone repository
git clone https://github.com/tmaths/crm.git
cd crm

# Create virtual environment
python -m venv venv

# Activate virtual environment
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# Install Python dependencies
pip install -r requirements.txt
```

#### 3. Environment Configuration

Create a `.env` file in the project root:

```env
# Django Settings
SECRET_KEY=your-super-secret-key-change-this-in-production
DEBUG=True
ALLOWED_HOSTS=localhost,127.0.0.1

# Database Configuration
DB_ENGINE=mssql
DB_NAME=dbexcel
DB_HOST=your-sql-server-hostname
DB_USER=your-domain\username
DB_PASSWORD=your-password
DB_PORT=1433

# Cache Configuration
CACHE_BACKEND=django.core.cache.backends.db.DatabaseCache
CACHE_LOCATION=dashboard_cache_table
CACHE_TIMEOUT=3600

# Celery Configuration
CELERY_BROKER_URL=redis://localhost:6379/0
CELERY_RESULT_BACKEND=redis://localhost:6379/0

# Logging
LOG_LEVEL=INFO
LOG_FILE=logs/application.log
```

#### 4. Database Setup

```bash
cd report

# Create database migrations
python manage.py makemigrations

# Apply migrations
python manage.py migrate

# Create cache table
python manage.py createcachetable

# Create superuser for admin access
python manage.py createsuperuser
```

#### 5. Run Development Server

```bash
python manage.py runserver
# Server available at http://localhost:8000
```

---

## Project Structure

### Directory Layout

```
crm/
├── report/                           # Django project root
│   ├── bulkrep/                      # Main application
│   │   ├── migrations/               # Database migrations
│   │   │   ├── 0001_initial.py
│   │   │   └── __init__.py
│   │   ├── templates/                # HTML templates
│   │   │   ├── base.html             # Base template
│   │   │   ├── dashboard.html        # Dashboard view
│   │   │   ├── single_report.html    # Single report form
│   │   │   └── bulk_report.html      # Bulk report form
│   │   ├── static/                   # Static files
│   │   │   ├── css/
│   │   │   │   ├── style.css
│   │   │   │   └── responsive.css
│   │   │   ├── js/
│   │   │   │   ├── dashboard.js      # Dashboard logic
│   │   │   │   ├── charts.js         # Chart initialization
│   │   │   │   └── filters.js        # Filter logic
│   │   │   └── images/
│   │   ├── models.py                 # Database models
│   │   ├── views.py                  # View logic
│   │   ├── urls.py                   # URL routing
│   │   ├── forms.py                  # Django forms
│   │   ├── serializers.py            # JSON serializers
│   │   ├── tasks.py                  # Celery tasks
│   │   └── admin.py                  # Django admin config
│   ├── accounts/                     # User management app
│   │   ├── models.py
│   │   ├── views.py
│   │   ├── urls.py
│   │   └── templates/
│   ├── report/                       # Project settings
│   │   ├── settings.py               # Django settings
│   │   ├── urls.py                   # Project URL config
│   │   ├── wsgi.py                   # WSGI config
│   │   └── celery.py                 # Celery config
│   ├── manage.py                     # Django management
│   └── logs/                         # Application logs
├── requirements.txt                  # Python dependencies
├── .env.example                      # Environment variables template
├── README.md                         # User documentation
├── TECHNICAL_DOCUMENTATION.md        # This file
└── .gitignore                        # Git ignore rules
```

### File Descriptions

| File/Folder | Purpose |
|-------------|---------|
| `bulkrep/models.py` | ORM models for Usagereport, SubscriberProductRate, ReportGeneration |
| `bulkrep/views.py` | Request handlers and business logic |
| `bulkrep/urls.py` | URL to view mapping |
| `bulkrep/forms.py` | Form validation and processing |
| `bulkrep/tasks.py` | Celery background tasks for report generation |
| `report/settings.py` | Django configuration, database, cache, middleware |
| `static/js/dashboard.js` | Client-side dashboard interactions |
| `static/js/charts.js` | Chart.js initialization and rendering |

---

## Database Schema

### Overview

The application uses **Microsoft SQL Server** with the following core models:

### Models

#### 1. **Usagereport**

Stores subscriber usage data for products.

```python
class Usagereport(models.Model):
    subscriber_id = models.IntegerField()
    subscriber_name = models.CharField(max_length=255)
    product_id = models.IntegerField()
    product_name = models.CharField(max_length=255)
    usage_quantity = models.DecimalField(max_digits=18, decimal_places=2)
    revenue = models.DecimalField(max_digits=18, decimal_places=2)
    date = models.DateField()
    created_at = models.DateTimeField(auto_now_add=True)
    
    class Meta:
        db_table = 'usagereport'
        indexes = [
            models.Index(fields=['subscriber_id', 'date']),
            models.Index(fields=['product_id']),
            models.Index(fields=['date']),
        ]
```

**Key Indexes:**
- `subscriber_id + date` (Composite) - For filtering by subscriber and date range
- `product_id` (Single) - For product-based queries
- `date` (Single) - For date range queries

#### 2. **SubscriberProductRate**

Pricing information for subscriber-product combinations.

```python
class SubscriberProductRate(models.Model):
    subscriber_id = models.IntegerField()
    product_id = models.IntegerField()
    rate_per_unit = models.DecimalField(max_digits=10, decimal_places=4)
    effective_date = models.DateField()
    
    class Meta:
        db_table = 'subscriber_product_rate'
        unique_together = ['subscriber_id', 'product_id', 'effective_date']
```

#### 3. **ReportGeneration**

Audit log for generated reports.

```python
class ReportGeneration(models.Model):
    STATUS_CHOICES = [
        ('pending', 'Pending'),
        ('processing', 'Processing'),
        ('completed', 'Completed'),
        ('failed', 'Failed'),
    ]
    
    user = models.ForeignKey(User, on_delete=models.PROTECT)
    report_type = models.CharField(max_length=50)  # 'single' or 'bulk'
    subscriber_ids = models.TextField()  # JSON list
    date_from = models.DateField()
    date_to = models.DateField()
    status = models.CharField(max_length=20, choices=STATUS_CHOICES)
    file_path = models.CharField(max_length=500, null=True, blank=True)
    error_message = models.TextField(null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    completed_at = models.DateTimeField(null=True, blank=True)
    
    class Meta:
        db_table = 'report_generation'
        indexes = [
            models.Index(fields=['user_id', 'created_at']),
            models.Index(fields=['status']),
        ]
```

### Database Queries

#### Query Optimization Examples

**Efficient Query with Indexes:**
```sql
-- Uses composite index on subscriber_id + date
SELECT SUM(revenue), COUNT(*) 
FROM usagereport 
WHERE subscriber_id = @subscriber_id 
  AND date BETWEEN @start_date AND @end_date
```

**Product Analysis:**
```sql
-- Uses index on product_id
SELECT TOP 10 product_name, SUM(revenue) as total_revenue
FROM usagereport
WHERE date BETWEEN @start_date AND @end_date
GROUP BY product_id, product_name
ORDER BY total_revenue DESC
```

---

## API Documentation

### Base URL

```
http://localhost:8000
```

### Authentication

All endpoints require Django session authentication or token-based authentication.

```javascript
// Example AJAX request with CSRF token
fetch('/api/endpoint/', {
    method: 'GET',
    headers: {
        'X-CSRFToken': getCookie('csrftoken'),
        'Content-Type': 'application/json'
    }
})
```

### Dashboard Endpoints

#### GET `/dashboard-api/`

Returns aggregated data for dashboard visualization.

**Query Parameters:**
```
?subscriber_id=1
?date_from=2024-01-01
?date_to=2024-12-31
```

**Response:**
```json
{
    "total_revenue": 150000,
    "total_usage": 5000,
    "subscriber_count": 25,
    "top_products": [
        {
            "product_id": 1,
            "product_name": "Product A",
            "revenue": 45000,
            "usage": 1500
        }
    ],
    "date_range": {
        "from": "2024-01-01",
        "to": "2024-12-31"
    }
}
```

#### GET `/api/usage-trends/`

Returns usage trend data with custom date filtering.

**Query Parameters:**
```
?period=monthly  # daily, weekly, monthly, quarterly, yearly
?subscriber_id=1 (optional)
?date_from=2024-01-01
?date_to=2024-12-31
```

**Response:**
```json
{
    "labels": ["Jan", "Feb", "Mar"],
    "datasets": [
        {
            "label": "Usage Trend",
            "data": [1000, 1200, 1100],
            "borderColor": "#36A2EB"
        }
    ]
}
```

#### GET `/api/new-subscribers-trend/`

Returns new subscriber acquisition trend.

**Response:**
```json
{
    "labels": ["2024-01", "2024-02", "2024-03"],
    "new_subscribers": [5, 8, 3],
    "cumulative_total": [100, 108, 111]
}
```

### Download Endpoints

#### GET `/download-churned-subscribers/`

Downloads CSV of churned subscribers.

**Response:** CSV file download

#### GET `/download-new-subscribers/`

Downloads CSV of new subscribers.

**Response:** CSV file download

### Report Generation Endpoints

#### POST `/api/generate-report/`

Initiates single or bulk report generation.

**Request Body:**
```json
{
    "report_type": "single",
    "subscriber_ids": [1, 2, 3],
    "date_from": "2024-01-01",
    "date_to": "2024-12-31",
    "include_bills": true,
    "include_products": true
}
```

**Response:**
```json
{
    "task_id": "abc123def456",
    "status": "processing",
    "estimated_time": 30
}
```

#### GET `/api/report-status/<task_id>/`

Checks status of report generation.

**Response:**
```json
{
    "task_id": "abc123def456",
    "status": "completed",
    "file_url": "/media/reports/report_2024_01_15.xlsx",
    "created_at": "2024-01-15T10:30:00Z"
}
```

---

## Frontend Architecture

### Template Structure

#### Base Template (`base.html`)

```html
<!DOCTYPE html>
<html>
<head>
    {% load static %}
    <title>{% block title %}Bulk Report Dashboard{% endblock %}</title>
    <link rel="stylesheet" href="{% static 'css/style.css' %}">
    <link rel="stylesheet" href="{% static 'css/responsive.css' %}">
    <script src="https://cdn.jsdelivr.net/npm/chart.js@3.9.1/dist/chart.min.js"></script>
</head>
<body>
    {% if user.is_authenticated %}
        <nav class="navbar">
            <a href="/dashboard/">Dashboard</a>
            <a href="/single-report/">Single Report</a>
            <a href="/bulk-report/">Bulk Report</a>
            <span>{{ user.username }} | <a href="/logout/">Logout</a></span>
        </nav>
    {% endif %}
    
    <main>
        {% block content %}{% endblock %}
    </main>
</body>
</html>
```

### JavaScript Architecture

#### Dashboard Module (`dashboard.js`)

```javascript
// Module pattern for dashboard functionality
const Dashboard = (() => {
    const state = {
        filters: {
            subscriber_id: null,
            date_from: null,
            date_to: null
        },
        data: {}
    };
    
    const init = () => {
        bindEvents();
        loadDashboardData();
    };
    
    const bindEvents = () => {
        document.querySelectorAll('[data-filter]').forEach(el => {
            el.addEventListener('change', handleFilterChange);
        });
    };
    
    const handleFilterChange = (e) => {
        const filterName = e.target.dataset.filter;
        state.filters[filterName] = e.target.value;
        loadDashboardData();
    };
    
    const loadDashboardData = () => {
        const params = new URLSearchParams(state.filters);
        fetch(`/dashboard-api/?${params}`)
            .then(response => response.json())
            .then(data => {
                state.data = data;
                renderCharts();
                renderMetrics();
            })
            .catch(error => console.error('Error loading data:', error));
    };
    
    const renderCharts = () => {
        // Chart rendering logic
    };
    
    const renderMetrics = () => {
        // Metrics display logic
    };
    
    return { init };
})();

// Initialize on page load
document.addEventListener('DOMContentLoaded', Dashboard.init);
```

#### Chart Module (`charts.js`)

```javascript
const ChartManager = (() => {
    const charts = {};
    
    const initializeChart = (canvasId, config) => {
        const ctx = document.getElementById(canvasId).getContext('2d');
        charts[canvasId] = new Chart(ctx, config);
    };
    
    const updateChart = (canvasId, newData) => {
        charts[canvasId].data.labels = newData.labels;
        charts[canvasId].data.datasets[0].data = newData.data;
        charts[canvasId].update();
    };
    
    const createRevenueTrendChart = (data) => {
        initializeChart('revenueTrendChart', {
            type: 'line',
            data: {
                labels: data.labels,
                datasets: [{
                    label: 'Revenue',
                    data: data.values,
                    borderColor: '#36A2EB',
                    backgroundColor: 'rgba(54, 162, 235, 0.1)',
                    tension: 0.1
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    title: { display: true, text: 'Revenue Trends' }
                }
            }
        });
    };
    
    return { initializeChart, updateChart, createRevenueTrendChart };
})();
```

### Responsive Design

CSS media queries for mobile, tablet, and desktop:

```css
/* Desktop (default) */
.dashboard-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
    gap: 20px;
}

/* Tablet */
@media (max-width: 768px) {
    .dashboard-grid {
        grid-template-columns: 1fr;
    }
    
    .navbar {
        flex-direction: column;
    }
}

/* Mobile */
@media (max-width: 480px) {
    .chart-container {
        max-height: 300px;
    }
    
    .filters {
        flex-direction: column;
    }
}
```

---

## Backend Implementation

### Views and URL Routing

#### URL Configuration (`urls.py`)

```python
from django.urls import path
from . import views

urlpatterns = [
    # Dashboard
    path('dashboard/', views.dashboard_view, name='dashboard'),
    path('dashboard-api/', views.dashboard_api, name='dashboard_api'),
    
    # Reports
    path('single-report/', views.single_report_view, name='single_report'),
    path('bulk-report/', views.bulk_report_view, name='bulk_report'),
    
    # API Endpoints
    path('api/usage-trends/', views.usage_trends_api, name='usage_trends'),
    path('api/new-subscribers-trend/', views.new_subscribers_trend_api, name='new_subscribers'),
    path('api/generate-report/', views.generate_report, name='generate_report'),
    path('api/report-status/<str:task_id>/', views.report_status, name='report_status'),
    
    # Downloads
    path('download-churned-subscribers/', views.download_churned_subscribers, name='download_churned'),
    path('download-new-subscribers/', views.download_new_subscribers, name='download_new'),
]
```

#### View Implementation (`views.py`)

```python
from django.shortcuts import render
from django.http import JsonResponse
from django.views.decorators.http import require_http_methods
from django.contrib.auth.decorators import login_required
from django.db.models import Sum, Count, F
from .models import Usagereport, ReportGeneration
from .tasks import generate_report_task
from datetime import datetime, timedelta
import json

@login_required
@require_http_methods(["GET"])
def dashboard_api(request):
    """Returns aggregated dashboard data"""
    # Get filters from request
    subscriber_id = request.GET.get('subscriber_id')
    date_from = request.GET.get('date_from', (datetime.now() - timedelta(days=30)).date())
    date_to = request.GET.get('date_to', datetime.now().date())
    
    # Build query
    query = Usagereport.objects.filter(date__range=[date_from, date_to])
    
    if subscriber_id:
        query = query.filter(subscriber_id=subscriber_id)
    
    # Aggregate data
    aggregates = query.aggregate(
        total_revenue=Sum('revenue'),
        total_usage=Sum('usage_quantity'),
        subscriber_count=Count('subscriber_id', distinct=True),
        product_count=Count('product_id', distinct=True)
    )
    
    # Get top products
    top_products = query.values('product_id', 'product_name').annotate(
        revenue=Sum('revenue'),
        usage=Sum('usage_quantity')
    ).order_by('-revenue')[:10]
    
    return JsonResponse({
        'total_revenue': float(aggregates['total_revenue'] or 0),
        'total_usage': float(aggregates['total_usage'] or 0),
        'subscriber_count': aggregates['subscriber_count'] or 0,
        'product_count': aggregates['product_count'] or 0,
        'top_products': list(top_products),
        'date_range': {
            'from': str(date_from),
            'to': str(date_to)
        }
    })

@login_required
@require_http_methods(["POST"])
def generate_report(request):
    """Initiate asynchronous report generation"""
    try:
        data = json.loads(request.body)
        
        # Validate input
        report_type = data.get('report_type')
        subscriber_ids = data.get('subscriber_ids', [])
        
        # Create tracking record
        report_gen = ReportGeneration.objects.create(
            user=request.user,
            report_type=report_type,
            subscriber_ids=json.dumps(subscriber_ids),
            date_from=data.get('date_from'),
            date_to=data.get('date_to'),
            status='pending'
        )
        
        # Queue Celery task
        task = generate_report_task.delay(report_gen.id)
        
        return JsonResponse({
            'task_id': task.id,
            'status': 'processing',
            'estimated_time': 30
        })
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=400)
```

### Model Methods

```python
class Usagereport(models.Model):
    # ... fields ...
    
    @classmethod
    def get_revenue_by_subscriber(cls, date_from, date_to):
        """Get revenue aggregated by subscriber"""
        return cls.objects.filter(
            date__range=[date_from, date_to]
        ).values('subscriber_id', 'subscriber_name').annotate(
            total_revenue=Sum('revenue')
        ).order_by('-total_revenue')
    
    @classmethod
    def get_top_products(cls, date_from, date_to, limit=10):
        """Get top products by revenue"""
        return cls.objects.filter(
            date__range=[date_from, date_to]
        ).values('product_id', 'product_name').annotate(
            total_revenue=Sum('revenue'),
            total_usage=Sum('usage_quantity')
        ).order_by('-total_revenue')[:limit]
```

---

## Caching Strategy

### Configuration

```python
# settings.py
CACHES = {
    'default': {
        'BACKEND': 'django.core.cache.backends.db.DatabaseCache',
        'LOCATION': 'dashboard_cache_table',
        'TIMEOUT': 3600,  # 1 hour
        'OPTIONS': {
            'MAX_ENTRIES': 10000,
            'CULL_FREQUENCY': 3,  # Remove 1/3 entries when full
        }
    }
}
```

### Cache Usage

```python
from django.core.cache import cache

# Cache dashboard data
cache_key = f"dashboard_{subscriber_id}_{date_from}_{date_to}"
cached_data = cache.get(cache_key)

if cached_data is None:
    # Expensive query
    data = Usagereport.objects.filter(...).values(...)
    cache.set(cache_key, data, timeout=3600)
else:
    data = cached_data

return JsonResponse(data)
```

### Cache Invalidation

```python
# Invalidate cache on data update
@receiver(post_save, sender=Usagereport)
def invalidate_dashboard_cache(sender, instance, **kwargs):
    """Clear relevant cache entries when usage data changes"""
    pattern = f"dashboard_{instance.subscriber_id}_*"
    # Clear cache logic
    cache.delete_pattern(pattern)
```

---

## Error Handling & Logging

### Logging Configuration

```python
# settings.py
LOGGING = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
        'verbose': {
            'format': '{levelname} {asctime} {module} {process:d} {thread:d} {message}',
            'style': '{',
        },
        'simple': {
            'format': '{levelname} {asctime} {message}',
            'style': '{',
        },
    },
    'filters': {
        'require_debug_false': {
            '()': 'django.utils.log.RequireDebugFalse',
        },
        'require_debug_true': {
            '()': 'django.utils.log.RequireDebugTrue',
        },
    },
    'handlers': {
        'console': {
            'level': 'INFO',
            'class': 'logging.StreamHandler',
            'formatter': 'simple',
        },
        'file': {
            'level': 'INFO',
            'class': 'logging.handlers.RotatingFileHandler',
            'filename': 'logs/application.log',
            'maxBytes': 1024 * 1024 * 10,  # 10MB
            'backupCount': 5,
            'formatter': 'verbose',
        },
    },
    'loggers': {
        'django': {
            'handlers': ['console', 'file'],
            'level': 'INFO',
            'propagate': True,
        },
        'bulkrep': {
            'handlers': ['console', 'file'],
            'level': 'DEBUG',
            'propagate': True,
        },
    },
}
```

### Error Handling Example

```python
import logging

logger = logging.getLogger(__name__)

def safe_report_generation(report_id):
    """Generate report with comprehensive error handling"""
    try:
        report = ReportGeneration.objects.get(id=report_id)
        report.status = 'processing'
        report.save()
        
        # Processing logic
        data = process_report_data(report)
        file_path = create_excel_file(data)
        
        report.status = 'completed'
        report.file_path = file_path
        report.completed_at = datetime.now()
        report.save()
        
        logger.info(f"Report {report_id} generated successfully")
        
    except Usagereport.DoesNotExist as e:
        logger.error(f"Report data not found for {report_id}: {str(e)}")
        report.status = 'failed'
        report.error_message = 'Required data not found'
        report.save()
        
    except Exception as e:
        logger.exception(f"Unexpected error generating report {report_id}")
        report.status = 'failed'
        report.error_message = str(e)
        report.save()
```

---

## Performance Optimization

### Database Query Optimization

#### N+1 Query Problem Prevention

```python
# Bad: N+1 queries
reports = Usagereport.objects.all()
for report in reports:
    print(report.subscriber_id)  # Extra query for each

# Good: Use select_related and prefetch_related
reports = Usagereport.objects.all().select_related('subscriber')
```

#### Query Analysis

```python
# Use Django Debug Toolbar
from django.db import connection
from django.test.utils import CaptureQueriesContext

with CaptureQueriesContext(connection) as context:
    # Your query here
    data = Usagereport.objects.filter(date__gte='2024-01-01').values(...)
    
print(f"Number of queries: {len(context.captured_queries)}")
for query in context.captured_queries:
    print(query['sql'])
    print(f"Time: {query['time']}")
```

### Bulk Operations

```python
# Bad: Individual saves in loop
for data in large_dataset:
    Usagereport.objects.create(**data)

# Good: Bulk create
Usagereport.objects.bulk_create(
    [Usagereport(**data) for data in large_dataset],
    batch_size=1000
)
```

### Connection Pooling

```python
# settings.py - SQL Server connection pool
DATABASES = {
    'default': {
        'ENGINE': 'mssql',
        'NAME': 'dbexcel',
        'OPTIONS': {
            'driver': 'ODBC Driver 17 for SQL Server',
            'Trusted_connection': 'yes',
            'Connection Pooling': 'true',
            'Min Pool Size': 5,
            'Max Pool Size': 10,
        },
    }
}
```

---

## Deployment Guide

### Production Checklist

- [ ] Set `DEBUG=False` in production
- [ ] Generate strong `SECRET_KEY`
- [ ] Configure allowed hosts
- [ ] Set up SSL/HTTPS
- [ ] Configure static file serving (Nginx/Apache)
- [ ] Set up Celery worker with Supervisor
- [ ] Configure log rotation
- [ ] Set up database backups
- [ ] Configure email settings for alerts
- [ ] Set up monitoring and alerting

### Deployment Steps

#### 1. Prepare Server

```bash
# Update system
apt-get update && apt-get upgrade -y

# Install Python and dependencies
apt-get install -y python3.8 python3-pip python3-venv
apt-get install -y msodbcsql17

# Create application directory
mkdir -p /var/www/crm
cd /var/www/crm
```

#### 2. Setup Application

```bash
# Clone repository
git clone https://github.com/tmaths/crm.git .

# Create virtual environment
python3 -m venv venv
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
pip install gunicorn  # WSGI server
```

#### 3. Configure Django

```bash
# Collect static files
python manage.py collectstatic --no-input

# Create superuser
python manage.py createsuperuser
```

#### 4. Setup Gunicorn

Create `/etc/systemd/system/crm-gunicorn.service`:

```ini
[Unit]
Description=Gunicorn Application Server for CRM
After=network.target

[Service]
User=www-data
Group=www-data
WorkingDirectory=/var/www/crm
ExecStart=/var/www/crm/venv/bin/gunicorn --workers 4 --bind unix:/var/www/crm/gunicorn.sock report.wsgi:application

[Install]
WantedBy=multi-user.target
```

Start service:
```bash
systemctl start crm-gunicorn
systemctl enable crm-gunicorn
```

#### 5. Setup Nginx

Create `/etc/nginx/sites-available/crm`:

```nginx
upstream crm_app {
    server unix:/var/www/crm/gunicorn.sock fail_timeout=0;
}

server {
    listen 80;
    server_name yourdomain.com;
    
    client_max_body_size 100M;
    
    location /static/ {
        alias /var/www/crm/staticfiles/;
    }
    
    location / {
        proxy_pass http://crm_app;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

Enable site:
```bash
ln -s /etc/nginx/sites-available/crm /etc/nginx/sites-enabled/
systemctl restart nginx
```

---

## Testing

### Unit Tests

```python
from django.test import TestCase
from bulkrep.models import Usagereport

class UsagereportTestCase(TestCase):
    def setUp(self):
        Usagereport.objects.create(
            subscriber_id=1,
            subscriber_name="Test Sub",
            product_id=1,
            product_name="Test Product",
            usage_quantity=100,
            revenue=1000,
            date='2024-01-01'
        )
    
    def test_get_revenue_by_subscriber(self):
        revenue = Usagereport.get_revenue_by_subscriber(
            '2024-01-01', '2024-01-31'
        )
        self.assertEqual(revenue[0]['total_revenue'], 1000)
```

### Integration Tests

```python
from django.test import Client, TestCase
from django.contrib.auth.models import User

class DashboardAPITestCase(TestCase):
    def setUp(self):
        self.client = Client()
        self.user = User.objects.create_user('testuser', 'test@example.com', 'pass')
    
    def test_dashboard_api_requires_login(self):
        response = self.client.get('/dashboard-api/')
        self.assertEqual(response.status_code, 302)  # Redirect to login
    
    def test_dashboard_api_returns_json(self):
        self.client.login(username='testuser', password='pass')
        response = self.client.get('/dashboard-api/')
        self.assertEqual(response.status_code, 200)
        self.assertIn('total_revenue', response.json())
```

### Running Tests

```bash
# Run all tests
python manage.py test

# Run specific test class
python manage.py test bulkrep.tests.UsagereportTestCase

# Run with coverage
pip install coverage
coverage run --source='.' manage.py test
coverage report
```

---

## Security Considerations

### Authentication & Authorization

- **Use HTTPS** in production to encrypt credentials
- **Implement CSRF protection** (enabled by default in Django)
- **Use Django's authentication system** for user management
- **Implement role-based access control** (RBAC)

```python
# Check permissions in views
from django.contrib.auth.decorators import permission_required

@permission_required('bulkrep.can_generate_reports')
def generate_report(request):
    # Only users with permission can access
    pass
```

### Database Security

- **Use parameterized queries** to prevent SQL injection (Django ORM does this)
- **Never commit `.env` files** with secrets
- **Use strong database passwords**
- **Restrict database user permissions** to minimum required

```python
# Bad: SQL injection vulnerable
query = f"SELECT * FROM users WHERE id = {user_id}"

# Good: Django ORM prevents injection
user = User.objects.get(id=user_id)
```

### File Upload Security

- **Validate file types** before processing
- **Store uploads outside web root**
- **Limit file sizes**
- **Scan for malware** (if applicable)

```python
def handle_file_upload(file):
    ALLOWED_TYPES = ['application/vnd.ms-excel', 'text/csv']
    MAX_SIZE = 10 * 1024 * 1024  # 10MB
    
    if file.size > MAX_SIZE:
        raise ValueError("File too large")
    
    if file.content_type not in ALLOWED_TYPES:
        raise ValueError("Invalid file type")
    
    # Process file
```

### API Security

- **Implement rate limiting** to prevent abuse
- **Use API keys/tokens** for external integrations
- **Validate all input** (Django forms do this)
- **Log suspicious activity**

```python
from django.views.decorators.cache import cache_page

@cache_page(60)  # Cache for 60 seconds
@login_required
def api_endpoint(request):
    # Rate limiting and caching applied
    pass
```

---

## Troubleshooting

### Common Issues

#### 1. Database Connection Errors

**Error:** `pyodbc.InterfaceError: ('IM002', '[IM002]')`

**Solution:**
```bash
# Verify ODBC driver installation
odbcad32.exe  # Windows

# Check connection string
python -c "import pyodbc; print(pyodbc.drivers())"

# Test connection
python manage.py dbshell
```

#### 2. Static Files Not Loading

**Error:** 404 for CSS/JS files in production

**Solution:**
```bash
# Collect static files
python manage.py collectstatic

# Verify Nginx/Apache serving from correct directory
# Check STATIC_ROOT and STATIC_URL in settings.py
```

#### 3. Celery Tasks Not Processing

**Error:** Tasks stay in pending state

**Solution:**
```bash
# Check Redis connection
redis-cli ping

# Verify Celery worker is running
celery -A report worker -l info

# Check task queue
redis-cli LLEN celery

# Inspect Celery
celery -A report inspect active
```

#### 4. Slow Dashboard Queries

**Issue:** Dashboard takes >5 seconds to load

**Solution:**
```python
# Enable query logging
LOGGING['loggers']['django.db.backends'] = {
    'level': 'DEBUG',
}

# Run query analysis
from django.db import connection
from django.test.utils import CaptureQueriesContext

# Check for missing indexes
# Monitor with: SELECT * FROM sys.dm_db_missing_index_details
```

#### 5. Memory Leaks in Report Generation

**Solution:**
```python
# Use generators for large datasets
def bulk_generate_reports(subscriber_ids):
    for sub_id in subscriber_ids:
        yield generate_single_report(sub_id)
        # Memory freed after each iteration

# Monitor memory usage
import psutil
process = psutil.Process()
print(process.memory_info().rss / 1024 / 1024)  # MB
```

### Debug Mode

Enable debug logging for troubleshooting:

```python
# settings.py
DEBUG = True
LOGGING['loggers']['bulkrep']['level'] = 'DEBUG'

# In code
import logging
logger = logging.getLogger(__name__)
logger.debug(f"Debug info: {variable}")
logger.info(f"Informational: {message}")
logger.warning(f"Warning: {issue}")
logger.error(f"Error: {exception}")
```

---

## Glossary

| Term | Definition |
|------|-----------|
| **ORM** | Object-Relational Mapping - abstraction layer for database operations |
| **WSGI** | Web Server Gateway Interface - standard for Python web applications |
| **ODBC** | Open Database Connectivity - standard for database connections |
| **CSRF** | Cross-Site Request Forgery - security token for form submissions |
| **Cache** | Temporary storage to reduce database queries |
| **Celery** | Distributed task queue for asynchronous processing |
| **Gunicorn** | WSGI HTTP Server for running Django applications |
| **Nginx** | Reverse proxy and web server |
| **Pagination** | Breaking large datasets into manageable chunks |
| **Query Optimization** | Improving database query performance |

---

## Additional Resources

- [Django Documentation](https://docs.djangoproject.com/)
- [SQL Server Documentation](https://docs.microsoft.com/en-us/sql/)
- [Chart.js Documentation](https://www.chartjs.org/docs/latest/)
- [OpenPyXL Documentation](https://openpyxl.readthedocs.io/)
- [Celery Documentation](https://docs.celeryproject.io/)

---

**Document Version:** 1.0  
**Last Updated:** 2026-06-15  
**Maintainer:** Development Team
