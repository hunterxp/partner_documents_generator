@echo off

echo "Installing dependencies..."
call ".\run_npm_install.bat"

echo "Checking for .env file..."
IF NOT EXIST .env (
    echo "Environment variables file (.env) not found"
)

echo "Checking for template.docx..."
IF NOT EXIST template.docx (
    echo "Template document (template.docx) not found"
)

echo "Running generation script..."
npm run generate

echo "Script completed. Waiting for user input to exit..."
pause