EPPlus Excel Report Generator
This project generates Excel reports using EPPlus and runs inside a Docker container. The application is built using .NET 6 and can be deployed in a Linux environment.

Prerequisites:
Docker installed on your machine
.NET 6 SDK installed locally (for local development)
Steps to Run the Application:
1. Build the Docker Image:
You can build the Docker image using the following command:


docker build -t epplus-excel-app .
2. Run the Docker Container:
Once the image is built, you can run the container and generate the Excel report using:


docker run --rm -v $(pwd)/output:/app/output epplus-excel-app
The -v $(pwd)/output:/app/output flag mounts the local output directory so that the generated Excel report is saved on your host machine.
3. Output:
The Excel file will be generated in the output folder in your current working directory on your host machine.

Dockerfile Explanation:
The application is built using a multi-stage Docker build.
The first stage (build-env) uses the .NET SDK to restore dependencies, build the project, and publish it to an output directory.
The second stage runs the app using the .NET runtime in a Linux environment.
Font dependencies (via fontconfig) are installed to ensure proper rendering of fonts in the Excel file.
Key Files:
Program.cs: Contains the logic to generate the Excel report using EPPlus.
Dockerfile: Contains the steps to build and run the app inside a Docker container.
Notes:
Ensure you have an output folder in your current directory before running the container, as the report will be saved in this folder.
If you encounter any errors related to paths, make sure that there are no hardcoded paths in your .csproj file.





