# Deployment Guide

This document provides comprehensive instructions for deploying DocStyleVerify.API in various environments.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Environment Configuration](#environment-configuration)
- [Local Development](#local-development)
- [Docker Deployment](#docker-deployment)
- [Production Deployment](#production-deployment)
- [Cloud Deployment](#cloud-deployment)
- [Performance Optimization](#performance-optimization)
- [Monitoring and Logging](#monitoring-and-logging)
- [Security Considerations](#security-considerations)
- [Troubleshooting](#troubleshooting)

## Prerequisites

### System Requirements
- **Operating System**: Windows 10+, macOS 10.15+, or Linux (Ubuntu 20.04+)
- **Runtime**: .NET 9.0 Runtime
- **Database**: PostgreSQL 12+
- **Memory**: Minimum 2GB RAM (4GB+ recommended for production)
- **Storage**: 10GB+ free space for application and templates
- **CPU**: 2+ cores recommended

### Software Dependencies
- .NET 9.0 SDK (for development)
- PostgreSQL 12 or higher
- Nginx (for production reverse proxy)
- Docker (optional, for containerized deployment)

## Environment Configuration

### Environment Variables

Create environment-specific configuration files or set the following environment variables:

```bash
# Database Configuration
CONNECTIONSTRINGS__DEFAULTCONNECTION="Host=localhost;Database=DocStyleVerifyDB;Username=postgres;Password=your_password"

# Application Settings
ASPNETCORE_ENVIRONMENT=Production
ASPNETCORE_URLS=http://+:5000;https://+:5001

# Template Storage
TEMPLATESTORAGE__PATH="/app/templates"

# Document Processing
DOCUMENTPROCESSING__MAXFILESIZEBYTES=104857600
DOCUMENTPROCESSING__PROCESSINGTIMEOUTMINUTES=10

# Logging
LOGGING__LOGLEVEL__DEFAULT=Information
LOGGING__LOGLEVEL__DOCSTYLEVERIFY_API=Debug

# Security
ALLOWEDHOSTS="localhost,your-domain.com"
```

### Configuration Files

#### appsettings.Production.json
```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Host=prod-db-server;Database=DocStyleVerifyDB;Username=app_user;Password=secure_password;SSL Mode=Require;"
  },
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning",
      "DocStyleVerify.API": "Information"
    },
    "Console": {
      "IncludeScopes": false
    },
    "File": {
      "Path": "/var/log/docstyleverify/app.log",
      "MinLevel": "Warning"
    }
  },
  "AllowedHosts": "your-domain.com,api.your-domain.com",
  "TemplateStorage": {
    "Path": "/var/lib/docstyleverify/templates"
  },
  "DocumentProcessing": {
    "MaxFileSizeBytes": 104857600,
    "ProcessingTimeoutMinutes": 15
  },
  "Kestrel": {
    "Limits": {
      "MaxRequestBodySize": 104857600
    }
  }
}
```

## Local Development

### Setup Steps

1. **Clone Repository**
   ```bash
   git clone <repository-url>
   cd DocStyleVerify.API
   ```

2. **Install Dependencies**
   ```bash
   dotnet restore
   ```

3. **Setup Database**
   ```bash
   # Create database
   createdb DocStyleVerifyDB
   
   # Run migrations
   dotnet ef database update
   ```

4. **Configure Settings**
   ```bash
   # Copy example settings
   cp appsettings.example.json appsettings.Development.json
   
   # Edit connection string and other settings
   nano appsettings.Development.json
   ```

5. **Run Application**
   ```bash
   dotnet run
   ```

The application will be available at `https://localhost:5001`.

## Docker Deployment

### Dockerfile
```dockerfile
FROM mcr.microsoft.com/dotnet/aspnet:9.0 AS base
WORKDIR /app
EXPOSE 80
EXPOSE 443

FROM mcr.microsoft.com/dotnet/sdk:9.0 AS build
WORKDIR /src
COPY ["DocStyleVerify.API.csproj", "."]
RUN dotnet restore "DocStyleVerify.API.csproj"
COPY . .
WORKDIR "/src"
RUN dotnet build "DocStyleVerify.API.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "DocStyleVerify.API.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .

# Create directories for templates and logs
RUN mkdir -p /app/templates /app/logs

# Set permissions
RUN chown -R app:app /app
USER app

ENTRYPOINT ["dotnet", "DocStyleVerify.API.dll"]
```

### Docker Compose
```yaml
version: '3.8'

services:
  docstyleverify-api:
    build: .
    ports:
      - "5000:80"
      - "5001:443"
    environment:
      - ASPNETCORE_ENVIRONMENT=Production
      - ASPNETCORE_URLS=http://+:80;https://+:443
      - ASPNETCORE_Kestrel__Certificates__Default__Password=your_cert_password
      - ASPNETCORE_Kestrel__Certificates__Default__Path=/https/aspnetapp.pfx
      - ConnectionStrings__DefaultConnection=Host=postgres;Database=DocStyleVerifyDB;Username=postgres;Password=postgres
    volumes:
      - ./templates:/app/templates
      - ./logs:/app/logs
      - ~/.aspnet/https:/https:ro
    depends_on:
      - postgres
    restart: unless-stopped

  postgres:
    image: postgres:15
    environment:
      - POSTGRES_DB=DocStyleVerifyDB
      - POSTGRES_USER=postgres
      - POSTGRES_PASSWORD=postgres
    volumes:
      - postgres_data:/var/lib/postgresql/data
      - ./init.sql:/docker-entrypoint-initdb.d/init.sql
    ports:
      - "5432:5432"
    restart: unless-stopped

  nginx:
    image: nginx:alpine
    ports:
      - "80:80"
      - "443:443"
    volumes:
      - ./nginx.conf:/etc/nginx/nginx.conf
      - ./ssl:/etc/nginx/ssl
    depends_on:
      - docstyleverify-api
    restart: unless-stopped

volumes:
  postgres_data:
```

### Build and Run with Docker
```bash
# Build the image
docker build -t docstyleverify-api .

# Run with Docker Compose
docker-compose up -d

# View logs
docker-compose logs -f docstyleverify-api

# Stop services
docker-compose down
```

## Production Deployment

### Server Setup (Ubuntu 20.04)

1. **Install Prerequisites**
   ```bash
   # Update system
   sudo apt update && sudo apt upgrade -y
   
   # Install .NET 9.0 Runtime
   wget https://packages.microsoft.com/config/ubuntu/20.04/packages-microsoft-prod.deb -O packages-microsoft-prod.deb
   sudo dpkg -i packages-microsoft-prod.deb
   sudo apt update
   sudo apt install -y aspnetcore-runtime-9.0
   
   # Install PostgreSQL
   sudo apt install -y postgresql postgresql-contrib
   
   # Install Nginx
   sudo apt install -y nginx
   ```

2. **Create Application User**
   ```bash
   sudo useradd -r -s /bin/false docstyleverify
   sudo mkdir -p /var/lib/docstyleverify
   sudo chown docstyleverify:docstyleverify /var/lib/docstyleverify
   ```

3. **Setup Database**
   ```bash
   sudo -u postgres psql
   CREATE DATABASE "DocStyleVerifyDB";
   CREATE USER docstyleverify WITH PASSWORD 'secure_password';
   GRANT ALL PRIVILEGES ON DATABASE "DocStyleVerifyDB" TO docstyleverify;
   \q
   ```

4. **Deploy Application**
   ```bash
   # Create application directory
   sudo mkdir -p /var/www/docstyleverify
   
   # Copy application files
   sudo cp -r ./publish/* /var/www/docstyleverify/
   
   # Set permissions
   sudo chown -R docstyleverify:docstyleverify /var/www/docstyleverify
   sudo chmod +x /var/www/docstyleverify/DocStyleVerify.API
   
   # Create directories
   sudo mkdir -p /var/lib/docstyleverify/templates
   sudo mkdir -p /var/log/docstyleverify
   sudo chown -R docstyleverify:docstyleverify /var/lib/docstyleverify
   sudo chown -R docstyleverify:docstyleverify /var/log/docstyleverify
   ```

### Systemd Service

Create `/etc/systemd/system/docstyleverify.service`:
```ini
[Unit]
Description=DocStyleVerify API
After=network.target

[Service]
Type=notify
User=docstyleverify
Group=docstyleverify
WorkingDirectory=/var/www/docstyleverify
ExecStart=/usr/bin/dotnet /var/www/docstyleverify/DocStyleVerify.API.dll
Restart=always
RestartSec=10
KillSignal=SIGINT
SyslogIdentifier=docstyleverify
Environment=ASPNETCORE_ENVIRONMENT=Production
Environment=DOTNET_PRINT_TELEMETRY_MESSAGE=false

[Install]
WantedBy=multi-user.target
```

Enable and start the service:
```bash
sudo systemctl daemon-reload
sudo systemctl enable docstyleverify
sudo systemctl start docstyleverify
sudo systemctl status docstyleverify
```

### Nginx Configuration

Create `/etc/nginx/sites-available/docstyleverify`:
```nginx
server {
    listen 80;
    server_name api.your-domain.com;
    return 301 https://$server_name$request_uri;
}

server {
    listen 443 ssl http2;
    server_name api.your-domain.com;

    ssl_certificate /etc/nginx/ssl/your-domain.crt;
    ssl_certificate_key /etc/nginx/ssl/your-domain.key;
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_ciphers HIGH:!aNULL:!MD5;

    client_max_body_size 100M;
    client_body_timeout 60;
    client_header_timeout 60;
    send_timeout 60;

    location / {
        proxy_pass http://localhost:5000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_cache_bypass $http_upgrade;
        proxy_connect_timeout 60;
        proxy_send_timeout 60;
        proxy_read_timeout 60;
    }

    # Health check endpoint
    location /health {
        access_log off;
        proxy_pass http://localhost:5000/health;
    }

    # Static files (if any)
    location /static/ {
        alias /var/www/docstyleverify/wwwroot/;
        expires 1y;
        add_header Cache-Control "public, immutable";
    }
}
```

Enable the site:
```bash
sudo ln -s /etc/nginx/sites-available/docstyleverify /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl reload nginx
```

## Cloud Deployment

### Azure App Service

1. **Create App Service**
   ```bash
   # Create resource group
   az group create --name DocStyleVerifyRG --location "East US"
   
   # Create App Service plan
   az appservice plan create --name DocStyleVerifyPlan --resource-group DocStyleVerifyRG --sku B1 --is-linux
   
   # Create web app
   az webapp create --resource-group DocStyleVerifyRG --plan DocStyleVerifyPlan --name docstyleverify-api --runtime "DOTNETCORE:9.0"
   ```

2. **Configure Application Settings**
   ```bash
   az webapp config appsettings set --resource-group DocStyleVerifyRG --name docstyleverify-api --settings \
     "ConnectionStrings__DefaultConnection=Server=your-server.postgres.database.azure.com;Database=DocStyleVerifyDB;Port=5432;User Id=username;Password=password;Ssl Mode=Require;" \
     "ASPNETCORE_ENVIRONMENT=Production" \
     "TemplateStorage__Path=/home/templates"
   ```

3. **Deploy Application**
   ```bash
   # Publish application
   dotnet publish -c Release -o ./publish
   
   # Create deployment package
   cd publish
   zip -r ../deploy.zip .
   
   # Deploy to Azure
   az webapp deployment source config-zip --resource-group DocStyleVerifyRG --name docstyleverify-api --src deploy.zip
   ```

### AWS Elastic Beanstalk

1. **Prepare Deployment Package**
   ```bash
   dotnet publish -c Release -o ./aws-deploy
   cd aws-deploy
   zip -r ../docstyleverify-api.zip .
   ```

2. **Deploy with EB CLI**
   ```bash
   eb init docstyleverify-api --platform "64bit Amazon Linux 2023 v3.0.1 running .NET 8" --region us-east-1
   eb create production-env
   eb deploy
   ```

### Google Cloud Platform

1. **Build and Deploy**
   ```bash
   # Build Docker image
   docker build -t gcr.io/your-project/docstyleverify-api .
   
   # Push to Container Registry
   docker push gcr.io/your-project/docstyleverify-api
   
   # Deploy to Cloud Run
   gcloud run deploy docstyleverify-api \
     --image gcr.io/your-project/docstyleverify-api \
     --platform managed \
     --region us-central1 \
     --allow-unauthenticated \
     --memory 2Gi \
     --cpu 2 \
     --max-instances 10
   ```

## Performance Optimization

### Database Optimization

1. **Connection Pooling**
   ```json
   {
     "ConnectionStrings": {
       "DefaultConnection": "Host=localhost;Database=DocStyleVerifyDB;Username=postgres;Password=password;Pooling=true;MinPoolSize=5;MaxPoolSize=20;"
     }
   }
   ```

2. **Database Indexes**
   ```sql
   -- Create indexes for frequently queried columns
   CREATE INDEX idx_templates_status ON templates(status);
   CREATE INDEX idx_templates_created_on ON templates(created_on);
   CREATE INDEX idx_textstyles_template_id ON textstyles(template_id);
   CREATE INDEX idx_textstyles_style_type ON textstyles(style_type);
   CREATE INDEX idx_verificationresults_template_id ON verificationresults(template_id);
   CREATE INDEX idx_verificationresults_verification_date ON verificationresults(verification_date);
   CREATE INDEX idx_mismatches_verification_result_id ON mismatches(verification_result_id);
   ```

### Application Optimization

1. **Memory Configuration**
   ```json
   {
     "Kestrel": {
       "Limits": {
         "MaxRequestBodySize": 104857600,
         "MaxConcurrentConnections": 100,
         "MaxConcurrentUpgradedConnections": 100
       }
     }
   }
   ```

2. **Response Caching**
   ```csharp
   // In Program.cs
   builder.Services.AddResponseCaching();
   builder.Services.AddMemoryCache();
   
   app.UseResponseCaching();
   ```

### File System Optimization

1. **Template Storage**
   ```bash
   # Use fast SSD storage for templates
   sudo mkdir -p /var/lib/docstyleverify/templates
   sudo mount -t tmpfs -o size=1G tmpfs /var/lib/docstyleverify/cache
   ```

## Monitoring and Logging

### Application Insights (Azure)
```json
{
  "ApplicationInsights": {
    "InstrumentationKey": "your-instrumentation-key"
  },
  "Logging": {
    "ApplicationInsights": {
      "LogLevel": {
        "Default": "Information"
      }
    }
  }
}
```

### Prometheus Metrics
```csharp
// Add to Program.cs
builder.Services.AddPrometheusMetrics();
app.UsePrometheusMetrics();
```

### Health Checks
```csharp
// Add to Program.cs
builder.Services.AddHealthChecks()
    .AddDbContextCheck<DocStyleVerifyDbContext>()
    .AddCheck("disk_space", new DiskSpaceHealthCheck())
    .AddCheck("memory", new MemoryHealthCheck());

app.MapHealthChecks("/health");
```

### Log Aggregation with Serilog
```json
{
  "Serilog": {
    "Using": ["Serilog.AspNetCore"],
    "MinimumLevel": "Information",
    "WriteTo": [
      {
        "Name": "Console"
      },
      {
        "Name": "File",
        "Args": {
          "path": "/var/log/docstyleverify/app-.log",
          "rollingInterval": "Day",
          "retainedFileCountLimit": 30
        }
      },
      {
        "Name": "Seq",
        "Args": {
          "serverUrl": "http://localhost:5341"
        }
      }
    ]
  }
}
```

## Security Considerations

### SSL/TLS Configuration
```json
{
  "Kestrel": {
    "Certificates": {
      "Default": {
        "Path": "/etc/ssl/certs/docstyleverify.pfx",
        "Password": "certificate_password"
      }
    }
  }
}
```

### Security Headers
```csharp
// Add to Program.cs
app.Use(async (context, next) =>
{
    context.Response.Headers.Add("X-Content-Type-Options", "nosniff");
    context.Response.Headers.Add("X-Frame-Options", "DENY");
    context.Response.Headers.Add("X-XSS-Protection", "1; mode=block");
    context.Response.Headers.Add("Strict-Transport-Security", "max-age=31536000; includeSubDomains");
    await next();
});
```

### File Upload Security
```csharp
// Validate file types and sizes
[RequestSizeLimit(104857600)] // 100MB
[RequestFormLimits(MultipartBodyLengthLimit = 104857600)]
public async Task<IActionResult> Upload(IFormFile file)
{
    // Validation logic
}
```

## Troubleshooting

### Common Issues

1. **Database Connection Issues**
   ```bash
   # Check database status
   sudo systemctl status postgresql
   
   # Test connection
   psql -h localhost -U postgres -d DocStyleVerifyDB
   
   # Check logs
   sudo tail -f /var/log/postgresql/postgresql-*.log
   ```

2. **Application Performance Issues**
   ```bash
   # Check application logs
   sudo journalctl -u docstyleverify -f
   
   # Monitor resource usage
   htop
   df -h
   free -h
   
   # Check database performance
   sudo -u postgres psql -c "SELECT * FROM pg_stat_activity;"
   ```

3. **File Upload Issues**
   ```bash
   # Check disk space
   df -h /var/lib/docstyleverify
   
   # Check permissions
   ls -la /var/lib/docstyleverify/templates
   
   # Check Nginx configuration
   sudo nginx -t
   sudo systemctl status nginx
   ```

### Log Analysis

1. **Application Logs**
   ```bash
   # View recent logs
   sudo journalctl -u docstyleverify --since "1 hour ago"
   
   # Follow logs in real-time
   sudo journalctl -u docstyleverify -f
   
   # Filter by log level
   sudo journalctl -u docstyleverify -p err
   ```

2. **Database Logs**
   ```bash
   # PostgreSQL logs
   sudo tail -f /var/log/postgresql/postgresql-*.log
   
   # Query performance
   sudo -u postgres psql -c "SELECT query, calls, total_time, mean_time FROM pg_stat_statements ORDER BY total_time DESC LIMIT 10;"
   ```

### Backup and Recovery

1. **Database Backup**
   ```bash
   # Create backup
   pg_dump -h localhost -U postgres DocStyleVerifyDB > backup_$(date +%Y%m%d_%H%M%S).sql
   
   # Automated backup script
   #!/bin/bash
   BACKUP_DIR="/var/backups/docstyleverify"
   DATE=$(date +%Y%m%d_%H%M%S)
   pg_dump -h localhost -U postgres DocStyleVerifyDB | gzip > $BACKUP_DIR/db_backup_$DATE.sql.gz
   find $BACKUP_DIR -name "*.sql.gz" -mtime +7 -delete
   ```

2. **Application Backup**
   ```bash
   # Backup templates and configuration
   tar -czf /var/backups/docstyleverify_templates_$(date +%Y%m%d).tar.gz /var/lib/docstyleverify/templates
   ```

### Disaster Recovery

1. **Database Restore**
   ```bash
   # Stop application
   sudo systemctl stop docstyleverify
   
   # Restore database
   dropdb DocStyleVerifyDB
   createdb DocStyleVerifyDB
   gunzip -c backup.sql.gz | psql -U postgres DocStyleVerifyDB
   
   # Start application
   sudo systemctl start docstyleverify
   ```

2. **Application Restore**
   ```bash
   # Restore templates
   sudo tar -xzf templates_backup.tar.gz -C /
   sudo chown -R docstyleverify:docstyleverify /var/lib/docstyleverify/templates
   ```

This deployment guide covers the most common deployment scenarios and provides comprehensive instructions for setting up DocStyleVerify.API in production environments. 