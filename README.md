# Image Converter API

An Express.js API that converts images to text using Google Cloud Vision API and exports to text or Excel formats.

## Vercel Deployment

This app is now configured for Vercel deployment. Follow these steps:

### 1. Environment Variables

In your Vercel dashboard, add the following environment variable:

- `GOOGLE_APPLICATION_CREDENTIALS`: Your Google Cloud service account JSON credentials as a string

Example:
```
GOOGLE_APPLICATION_CREDENTIALS={"type":"service_account","project_id":"your-project-id","private_key_id":"...","private_key":"...","client_email":"...","client_id":"...","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_x509_cert_url":"..."}
```

### 2. Deploy to Vercel

1. Push your code to GitHub
2. Connect your repository to Vercel
3. Vercel will automatically detect the configuration and deploy

Or use Vercel CLI:
```bash
npm i -g vercel
vercel
```

### 3. API Endpoints

- `POST /api/convert` - Convert image to text/excel
- `POST /api/download` - Download converted files

## Local Development

```bash
npm install
NODE_ENV=development npm start
```

For local development, you can still use the keyFilename approach or set up the environment variable locally. 