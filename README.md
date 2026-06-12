# Bulk Report Dashboard

A comprehensive Django-based reporting and analytics dashboard for subscriber usage data and business intelligence.

## Overview

The Bulk Report Dashboard is a web application designed to generate detailed reports and provide analytics insights for subscriber usage data. It features an interactive dashboard with charts, filtering capabilities, and automated report generation for both individual subscribers and bulk operations.

## Features

### 📊 Interactive Dashboard
- **Real-time Analytics**: Interactive charts showing revenue trends, usage patterns, and subscriber metrics
- **Advanced Filtering**: Filter data by subscriber, date ranges, and custom time periods
- **Data Export**: Download charts and data as CSV files
- **Responsive Design**: Modern, mobile-friendly interface

### 📈 Key Metrics & Visualizations
- Revenue trends over time with subscriber filtering
- Top products by usage and revenue
- Top subscribers by activity
- Usage trends with custom date range support
- New subscriber acquisition trends
- Churn analysis and reporting

### 📋 Report Generation
- **Single Reports**: Generate detailed reports for individual subscribers
- **Bulk Reports**: Process multiple subscribers simultaneously
- **Excel Integration**: Automated Excel report generation with custom formatting
- **Date Range Selection**: Flexible date filtering for all reports
- **Product Billing**: Include/exclude product billing information

### 👥 User Management
- User authentication and authorization
- User impersonation for administrative purposes
- Activity tracking and audit logs
- Role-based access control

### 🔧 Technical Features
- **Database Caching**: Optimized performance with database-level caching
- **API Endpoints**: RESTful APIs for data retrieval and integration
- **Background Processing**: Efficient handling of large report generation tasks
- **Error Handling**: Comprehensive error logging and user feedback

## Technology Stack

- **Backend**: Django 2.1.15
- **Database**: Microsoft SQL Server with ODBC connectivity
- **Frontend**: HTML5, CSS3, JavaScript (Chart.js for visualizations)
- **Excel Processing**: OpenPyXL for advanced Excel manipulation
- **Caching**: Django database cache
- **Authentication**: Django built-in authentication system
- **Task Processing**: Celery for background tasks

## Installation

### Prerequisites

- Python 3.8+
- Microsoft SQL Server
- ODBC Driver 17 for SQL Server
- Git

### Setup Instructions

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd bulkreport
   ```

2. **Create virtual environment**
   ```bash
   python -m venv venv
   venv\Scripts\activate  # On Windows
   # source venv/bin/activate  # On Linux/Mac
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Environment Configuration**
   Create a `.env` file in the project root:
   ```env
   SECRET_KEY=your-secret-key-here
   DEBUG=True
   DB_ENGINE=mssql
   DB_NAME=dbexcel
   DB_HOST=your-sql-server-host
   DB_USER=your-username
   DB_PASSWORD=your-password
   DB_PORT=1433
   ```

5. **Database Setup**
   ```bash
   cd report
   python manage.py makemigrations
   python manage.py migrate
   python manage.py createcachetable
   ```

6. **Create Superuser**
   ```bash
   python manage.py createsuperuser
   ```

7. **Run Development Server**
   ```bash
   python manage.py runserver
   ```

   Access the application at `http://localhost:8000`

## Usage

### Dashboard Access
1. Navigate to `/dashboard/` for the main analytics dashboard
2. Use filters to customize data views:
   - **Subscriber Filter**: Filter all charts by specific subscriber
   - **Time Filters**: Select predefined periods or custom date ranges
   - **Global Filters**: Apply filters across all dashboard components

### Report Generation

#### Single Reports
1. Go to `/single-report/`
2. Select subscriber and date range
3. Choose report options (include bills, products)
4. Generate and download Excel report

#### Bulk Reports
1. Navigate to `/bulk-report/`
2. Select multiple subscribers or date ranges
3. Configure bulk processing options
4. Monitor progress and download completed reports

### Data Export
- **CSV Downloads**: Available for all chart data
- **Excel Reports**: Comprehensive formatted reports
- **Filtered Exports**: Maintain applied filters in exported data

## API Endpoints

### Dashboard API
- `GET /dashboard-api/` - Main dashboard data
- `GET /api/usage-trends/` - Usage trends with date filtering
- `GET /api/new-subscribers-trend/` - New subscriber acquisition data

### Download Endpoints
- `GET /download-churned-subscribers/` - Churned subscribers CSV
- `GET /download-new-subscribers/` - New subscribers CSV

## Database Schema

### Core Models
- **Usagereport**: Main usage data with subscriber, product, and date information
- **SubscriberProductRate**: Pricing information for subscriber-product combinations
- **ReportGeneration**: Tracking and audit log for generated reports

### Key Fields
- Subscriber information and identification
- Product usage and billing data
- Date-based filtering and reporting
- User activity and report generation tracking

## Configuration

### Cache Settings
The application uses database-level caching for improved performance:
```python
CACHES = {
    'default': {
        'BACKEND': 'django.core.cache.backends.db.DatabaseCache',
        'LOCATION': 'dashboard_cache_table',
        'TIMEOUT': 3600,
    }
}
```

### Database Configuration
Configured for Microsoft SQL Server with Windows Authentication:
```python
DATABASES = {
    'default': {
        'ENGINE': 'mssql',
        'NAME': 'dbexcel',
        'OPTIONS': {
            'driver': 'ODBC Driver 17 for SQL Server',
            'Trusted_connection': 'yes',
        },
    }
}
```

## Development

### Project Structure
```
bulkreport/
├── report/                 # Main Django project
│   ├── bulkrep/           # Core application
│   │   ├── models.py      # Data models
│   │   ├── views.py       # Business logic
│   │   ├── urls.py        # URL routing
│   │   └── templates/     # HTML templates
│   ├── accounts/          # User management
│   ├── report/            # Project settings
│   └── manage.py          # Django management
├── requirements.txt       # Python dependencies
└── README.md             # This file
```

### Key Components
- **Dashboard Views**: Interactive analytics and visualizations
- **Report Generation**: Excel processing and bulk operations
- **API Layer**: RESTful endpoints for data access
- **User Management**: Authentication and authorization
- **Caching Layer**: Performance optimization

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Commit changes (`git commit -am 'Add new feature'`)
4. Push to branch (`git push origin feature/new-feature`)
5. Create Pull Request

## License

This project is proprietary software. All rights reserved.

## Support

For technical support or questions:
- Check the Django documentation for framework-related issues
- Review the application logs for debugging information
- Contact the development team for application-specific support

---

**Note**: This application requires proper database configuration and ODBC driver installation for Microsoft SQL Server connectivity.