import mlflow
from mlflow.tracking import MlflowClient
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from dotenv import load_dotenv
import os

# Load environment variables from .env file
load_dotenv()

# Set MLflow Tracking URI from the .env file
tracking_uri = os.getenv("MLFLOW_TRACKING_URI")
if not tracking_uri:
    raise ValueError("MLFLOW_TRACKING_URI is not set in the .env file.")
mlflow.set_tracking_uri(tracking_uri)

client = MlflowClient()

def fetch_full_metric_history(run_id, metric_name):
    """
    Fetch full history of a given metric for a run.
    """
    try:
        history = client.get_metric_history(run_id, metric_name)
        return [(point.timestamp, point.value) for point in history]
    except Exception:
        return []

def fetch_all_experiment_metrics():
    """
    Fetch detailed metrics and metadata for all experiments.
    """
    experiments = mlflow.search_experiments()
    all_data = []

    for experiment in experiments:
        runs = client.search_runs(experiment_ids=[experiment.experiment_id])

        for run in runs:
            gpu_metric_histories = {}

            for i in range(12):  # Assuming up to 12 GPUs
                metric_name = f"system/gpu_{i}_memory_usage_percentage"
                history = fetch_full_metric_history(run.info.run_id, metric_name)
                if history:
                    # Convert history to a readable string
                    history_str = "; ".join(
                        [f"Timestamp: {ts}, Value: {val}" for ts, val in history]
                    )
                    gpu_metric_histories[f"GPU_{i}_History"] = history_str

            run_data = {
                "Experiment Name": experiment.name,
                "Run ID": run.info.run_id,
                "Run Name": run.info.run_name,
                "User": run.data.tags.get("mlflow.user", "Unknown"),
                "Source": run.data.tags.get("mlflow.source.name", "Unknown"),
                "Status": run.info.status,
                "Start Time": human_readable_date(run.info.start_time),
                "End Time": human_readable_date(run.info.end_time),
                "Duration (s)": (run.info.end_time - run.info.start_time) / 1000 if run.info.end_time else None,
                "Logged Models": ", ".join([artifact.path for artifact in client.list_artifacts(run.info.run_id)]),
                "Registered Models": run.data.tags.get("mlflow.log-model.history", "None"),
            }

            run_data.update(gpu_metric_histories)
            all_data.append(run_data)

    return pd.DataFrame(all_data)

def fetch_registered_models():
    """
    Fetch details of registered models and their versions.
    """
    try:
        registered_models = client.search_registered_models()
        all_model_data = []

        for model in registered_models:
            for version in model.latest_versions:
                all_model_data.append({
                    "Model Name": model.name,
                    "Version": version.version,
                    "Stage": version.current_stage,
                    "Source": version.source,
                    "Git Commit": version.tags.get("mlflow.source.git.commit", "No"),
                    "Creation Time": human_readable_date(version.creation_timestamp),
                })

        return pd.DataFrame(all_model_data)
    except Exception as e:
        print(f"Error fetching registered models: {e}")
        return pd.DataFrame()  # Return an empty DataFrame if no models are found or an error occurs

def human_readable_date(timestamp):
    """
    Convert timestamp to a human-readable format.
    """
    if timestamp is None:
        return None
    return datetime.fromtimestamp(timestamp / 1000).strftime("%d/%m/%Y %H:%M:%S")

def generate_excel(all_experiment_data, registered_models, output_file="experiment_metrics_summary.xlsx"):
    """
    Generate an Excel file with experiment metrics and registered model details.
    """
    wb = Workbook()

    # Experiment Metrics Sheet
    metrics_ws = wb.active
    metrics_ws.title = "Experiment Metrics"

    for row in dataframe_to_rows(all_experiment_data, index=False, header=True):
        metrics_ws.append(row)

    # Registered Models Sheet
    if not registered_models.empty:
        models_ws = wb.create_sheet(title="Registered Models")
        for row in dataframe_to_rows(registered_models, index=False, header=True):
            models_ws.append(row)

        # Summary about models
        total_registered_models = len(registered_models["Model Name"].unique())
        models_with_git_commit = len(registered_models[registered_models["Git Commit"] != "No"])
        print(f"Total Registered Models: {total_registered_models}")
        print(f"Models Linked to Git Commits: {models_with_git_commit}")

    else:
        print("No registered models found. Skipping Registered Models sheet.")

    wb.save(output_file)
    print(f"Saved metrics and models to {output_file}")

if __name__ == "__main__":
    experiment_data = fetch_all_experiment_metrics()
    registered_model_data = fetch_registered_models()

    if not experiment_data.empty:
        generate_excel(experiment_data, registered_model_data)
    else:
        print("No experiment data found.")
