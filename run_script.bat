:: Build the Docker image for pears_partnerships_data_entry.py
docker build -t il_fcs/pears_partnerships_data_entry:latest .
:: Create and start the Docker container
docker run --name pears_partnerships_data_entry il_fcs/pears_partnerships_data_entry:latest
:: Copy /example_outputs from the container to the build context
docker cp pears_partnerships_data_entry:/pears_partnerships_data_entry/example_outputs/ ./
:: Remove the container
docker rm pears_partnerships_data_entry
pause