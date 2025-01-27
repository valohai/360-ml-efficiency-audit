import os
import mlflow
from mlflow.tracking import MlflowClient
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from dotenv import load_dotenv

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


def calculate_gpu_utilization_and_history(run_id):
    """
    Calculate GPU utilization history and average utilization for all GPUs in a run.
    """
    gpu_histories = {}
    gpu_utilization = []
    gpu_utilization_per_gpu = {}

    for i in range(12):  # Assuming up to 12 GPUs
        metric_name = f"system/gpu_{i}_utilization_percentage"
        history = fetch_full_metric_history(run_id, metric_name)
        if history:
            # Store history as a readable string
            history_str = "; ".join([f"Timestamp: {ts}, Value: {val}" for ts, val in history])
            gpu_histories[f"GPU_{i}_History"] = history_str

            # Calculate average utilization for this GPU
            avg_utilization = sum([point[1] for point in history]) / len(history)
            gpu_utilization.append(avg_utilization)
            gpu_utilization_per_gpu[f"GPU_{i}_Average_Utilization (%)"] = avg_utilization

    # Calculate overall average GPU utilization
    overall_avg_utilization = sum(gpu_utilization) / len(gpu_utilization) if gpu_utilization else None

    return gpu_histories, gpu_utilization_per_gpu, overall_avg_utilization


def fetch_all_experiment_metrics():
    """
    Fetch detailed metrics, parameters, and metadata for all experiments.
    """
    experiments = mlflow.search_experiments()
    all_data = []

    for experiment in experiments:
        runs = client.search_runs(experiment_ids=[experiment.experiment_id])

        for run in runs:
            # Fetch GPU utilization history and averages
            gpu_histories, gpu_utilization_per_gpu, overall_gpu_utilization = calculate_gpu_utilization_and_history(
                run.info.run_id
            )

            # Check connection to Git, Dataset, or Versioned Environment
            git_commit = run.data.tags.get("mlflow.source.git.commit", "No")
            dataset = run.data.tags.get("mlflow.data.dataset", "No")
            versioned_env = (
                "Conda" if "mlflow.conda_env" in run.data.tags
                else "Requirements" if "mlflow.requirements" in run.data.tags
                else "Docker" if "mlflow.docker" in run.data.tags
                else "No"
            )

            # Extract parameters and concatenate them into a single string
            params = run.data.params
            parameters_str = ", ".join([f"{key}: {value}" for key, value in params.items()])

            # Prepare run data
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
                "Git Commit": git_commit,
                "Dataset": dataset,
                "Versioned Environment": versioned_env,
                "Parameters": parameters_str,
                "Average GPU Utilization (%)": overall_gpu_utilization,
            }

            # Add GPU histories and per-GPU utilization
            run_data.update(gpu_histories)
            run_data.update(gpu_utilization_per_gpu)

            all_data.append(run_data)

    return pd.DataFrame(all_data)


def human_readable_date(timestamp):
    """
    Convert timestamp to a human-readable format.
    """
    if timestamp is None:
        return None
    return datetime.fromtimestamp(timestamp / 1000).strftime("%d/%m/%Y %H:%M:%S")


def generate_excel(all_experiment_data, output_file="experiment_metrics_summary.xlsx"):
    """
    Generate an Excel file with experiment metrics.
    """
    wb = Workbook()

    # Experiment Metrics Sheet
    metrics_ws = wb.active
    metrics_ws.title = "Experiment Metrics"

    for row in dataframe_to_rows(all_experiment_data, index=False, header=True):
        metrics_ws.append(row)

    wb.save(output_file)
    print(f"Saved metrics to {output_file}")


if __name__ == "__main__":
    experiment_data = fetch_all_experiment_metrics()

    if not experiment_data.empty:
        generate_excel(experiment_data)
    else:
        print("No experiment data found.")
