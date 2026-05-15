# ICFGlueUpSync — Cloud Run container
# Runs ICFGlueUpSync.py and all utility scripts.
# Authentication via Application Default Credentials (service account
# assigned to the Cloud Run job — no key file required).

FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install dependencies first (layer-cached unless requirements.txt changes)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy all scripts
COPY ICFGlueUpSync.py .
COPY GlobalGlueUpSync.py .
COPY ExGetRecord.py .
COPY GetToken.py .
COPY GetFieldCodes.py .
COPY GetMultiMembers.py .
COPY DeleteMultiMembers.py .
COPY ContactAudit.py .

# Output files are written to /app — Cloud Run jobs are ephemeral, so
# output files are uploaded to Google Drive before the container exits.
# If a run fails mid-way, re-run from the Cloud Run console.

# Environment variables set at deploy time (not here):
#   GLUEUP_USE_SECRET_MANAGER=1
#   GOOGLE_CLOUD_PROJECT=gg-sync-492000

# Default command: run the full sync using GlobalGlueUpSync.py.
# Override at invocation time to run a utility script instead, e.g.:
#   python ExGetRecord.py --email someone@example.com
CMD ["python", "GlobalGlueUpSync.py"]