Deployment options for `gdp-dashboard` (make the app publicly available 24/7)

1) Render (recommended, simple)
- Create a Render Web Service using Docker or Private Docker and connect it to this GitHub repo.
- If Render asks for a build command and start command:
  - Build command: leave blank when using the included Dockerfile
  - Start command: `streamlit run app.py --server.port 8502 --server.address 0.0.0.0 --server.enableCORS false`
- Set `PORT` to `8502` if Render asks for a service port.
- Create two GitHub repository secrets:
  - `RENDER_API_KEY` - from Render account settings
  - `RENDER_SERVICE_ID` - the Render service ID for your Web Service
- Push to `main` and the workflow will build the image, push to GHCR, then update the Render service image.
- If you want to deploy manually, run `./render-deploy.sh` with `RENDER_API_KEY` and `RENDER_SERVICE_ID` set.
- The workflow also pushes the image to GitHub Container Registry at `ghcr.io/<owner>/gdp-dashboard:latest` and `ghcr.io/<owner>/gdp-dashboard:<sha>`.

2) Fly.io
- Install `flyctl` and create an app: `fly launch` and follow prompts.
- Build and deploy with `fly deploy` (the Dockerfile included works).

3) Docker on a VPS / Cloud VM
- Build the image and run it with:
```bash
docker build -t gdp-dashboard:latest .
docker run -d --restart unless-stopped -p 8502:8502 gdp-dashboard:latest
```
- Use a reverse proxy (Nginx) and TLS (Let's Encrypt) to expose a public link.

4) Streamlit Cloud
- Connect your GitHub repo to Streamlit Community Cloud and set `app.py` as the entry file. Streamlit Cloud runs the app continuously but may sleep on free plans.

5) Using the provided systemd unit
- Edit `deploy/streamlit.service` to set correct paths and `User`.
- Copy to `/etc/systemd/system/streamlit.service` and enable:
```bash
sudo cp deploy/streamlit.service /etc/systemd/system/streamlit.service
sudo systemctl daemon-reload
sudo systemctl enable --now streamlit.service
sudo journalctl -u streamlit -f
```

Notes
- For global availability and resilience, prefer a cloud-hosted service (Render, Fly.io, Docker on AWS/GCP/DigitalOcean) with HTTPS and a domain.
- The GitHub Actions workflow builds and pushes images to GHCR. Add `RENDER_API_KEY` and `RENDER_SERVICE_ID` in repo secrets to enable automatic Render deployment.
