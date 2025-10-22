# DM2-All-Projects-Reports

A comprehensive project management and reporting system with role-based access control.

## Features

- **Three-Portal System**: Admin, Manager, and TL & Coordinator portals
- **Real-time Data Integration**: Google Sheets and Excel file upload
- **Advanced Analytics**: Utilisation, Stack Ranking, and Team Performance calculations
- **3D Interactive UI**: Modern design with animations and visualizations
- **Feedback System**: Communication between managers and TLs/Coordinators

## Login Credentials

- **Admin**: `admin@dm2.com` / `admin123`
- **Manager**: `manager@dm2.com` / `manager123`
- **TL**: `tl@dm2.com` / `tl123`

## Data Sources

- Google Sheets integration with custom URL input
- Excel file upload (.xlsx) with automatic column mapping
- Real-time calculations with your exact formulas

## Live Demo

Visit the live application at: [Your Netlify URL will appear here]

**Deployment Status:** ✅ Ready for deployment
- Site ID: `222c95ac-1bb0-49a7-b747-6bde845b54d3`
- Deployment package: `dm2-project.zip` (23.8 KB)

## Technical Stack

- HTML5, CSS3, JavaScript (ES6+)
- Chart.js for data visualizations
- SheetJS for Excel parsing
- LocalStorage for data persistence
- Netlify for hosting

## Deployment

This application can be deployed on Render with automatic builds from the main branch.

### Deployment Instructions:

1. **GitHub Setup:**
   - Create a new repository on GitHub
   - Upload all project files to the repository
   - Make sure the repository is public

2. **Render Deployment:**
   - Go to [render.com](https://render.com)
   - Sign up/Login with your GitHub account
   - Click "New +" → "Web Service"
   - Connect your GitHub repository
   - Use these settings:
     - **Build Command**: `echo "Static site - no build required"`
     - **Start Command**: `npx serve -s . -l $PORT`
     - **Plan**: Free
   - Click "Create Web Service"

3. **Alternative Manual Upload:**
   - Use the `dm2-project.zip` file created earlier
   - Upload directly to Render or any static hosting service
