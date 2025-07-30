#!/bin/bash

# Style Suggestions Word Add-in Startup Script

echo "🚀 Starting Style Suggestions Word Add-in Development Environment"
echo "=================================================================="

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo "❌ Node.js is not installed. Please install Node.js (version 14 or higher) first."
    exit 1
fi

# Check Node.js version
NODE_VERSION=$(node -v | cut -d'v' -f2 | cut -d'.' -f1)
if [ "$NODE_VERSION" -lt 14 ]; then
    echo "❌ Node.js version 14 or higher is required. Current version: $(node -v)"
    exit 1
fi

echo "✅ Node.js version: $(node -v)"

# Check if we're in the right directory
if [ ! -f "package.json" ]; then
    echo "❌ Please run this script from the word-addin directory"
    exit 1
fi

# Install dependencies if node_modules doesn't exist
if [ ! -d "node_modules" ]; then
    echo "📦 Installing dependencies..."
    npm install
    
    if [ $? -ne 0 ]; then
        echo "❌ Failed to install dependencies"
        exit 1
    fi
    
    echo "✅ Dependencies installed successfully"
else
    echo "✅ Dependencies already installed"
fi

# Validate manifest
echo "🔍 Validating manifest..."
npm run validate

if [ $? -ne 0 ]; then
    echo "⚠️  Manifest validation failed, but continuing..."
fi

# Start development server in background
echo "🌐 Starting development server..."
npm run dev-server &
DEV_SERVER_PID=$!

# Wait a moment for server to start
sleep 3

# Check if server is running
if curl -s http://localhost:3000 > /dev/null; then
    echo "✅ Development server is running at https://localhost:3000"
else
    echo "❌ Development server failed to start"
    kill $DEV_SERVER_PID 2>/dev/null
    exit 1
fi

echo ""
echo "🎉 Setup Complete!"
echo "=================="
echo ""
echo "Next steps:"
echo "1. Open Microsoft Word"
echo "2. In another terminal, run: npm start"
echo "3. This will sideload the add-in into Word"
echo ""
echo "Development URLs:"
echo "- Task Pane: https://localhost:3000/taskpane.html"
echo "- Commands: https://localhost:3000/commands.html"
echo ""
echo "Useful commands:"
echo "- npm start        : Sideload the add-in"
echo "- npm stop         : Remove the add-in"
echo "- npm run build    : Build for production"
echo "- npm run validate : Validate manifest"
echo ""
echo "To stop the development server, press Ctrl+C"
echo ""

# Keep the script running to maintain the dev server
wait $DEV_SERVER_PID 