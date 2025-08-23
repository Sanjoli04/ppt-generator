# Use an official Python runtime as a parent image
FROM python:3.11-slim

# Set the working directory in the container
WORKDIR /app

# Copy the requirements file into the container
COPY requirements.txt .

# Install any needed packages specified in requirements.txt
# --no-cache-dir ensures we don't store unnecessary files, keeping the image small
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of your application code into the container
COPY . .

# Make port 80 available to the world outside this container
# Render will automatically use this port
EXPOSE 80

# Define the command to run your app using Gunicorn
# This tells Gunicorn to run the 'app' object from your 'main.py' file
CMD ["gunicorn", "--bind", "0.0.0.0:80", "main:app"]
