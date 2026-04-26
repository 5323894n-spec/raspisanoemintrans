<div align="center">
<img width="1200" height="475" alt="GHBanner" src="https://github.com/user-attachments/assets/0aa67016-6eaf-458a-adb2-6e31a0763ed6" />
</div>

# Run and deploy your AI Studio app

This contains everything you need to run your app locally.

View your app in AI Studio: https://ai.studio/apps/c1cc4b5d-148d-4ae5-ba9d-91daef4d0aab

## Run Locally

**Prerequisites:**  Node.js


1. Install dependencies:
   `npm install`
2. Set the `GEMINI_API_KEY` in [.env.local](.env.local) to your Gemini API key
3. Run the app:
   `npm run dev`

### Deployment to GitHub Pages
The project is configured with `base: '/Raspisanie-M/'` in `vite.config.ts`.
To deploy:
1. Run `npm run build`
2. Follow the GitHub Pages instructions for the `dist` folder.
