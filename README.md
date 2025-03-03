python app.py
gclound builds submit --tag gcr.io/gantt-451208/dash_gantt --project==gantt-451208

gcloud builds submit --tag gcr.io/gantt-451208/dash_gantt --project==gantt-451208


gcloud run deploy --image gcr.io/gantt-451208/dash_gantt --platform managed --project==gantt-451208 --allow-unauthenticated

gcloud run deploy --image gcr.io/gantt-451208/dash_gantt --platform managed --region asia-east1 --allow-unauthenticated
