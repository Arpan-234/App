# 📊 AI Excel & PowerPoint Automation Agent

![Streamlit](https://img.shields.io/badge/Streamlit-1.34-brightgreen)
![Python](https://img.shields.io/badge/Python-3.10+-blue)
![Claude AI](https://img.shields.io/badge/Claude%20AI-3.5%20Sonnet-purple)

An intelligent Streamlit application that automates Excel data analysis and PowerPoint presentation creation using Claude AI. Process, analyze, and visualize your data, then automatically generate professional presentations.

## 🎯 Features

### Excel Automation
- **📈 Data Cleaning**: Auto-detect and suggest data cleaning steps
- **🔍 Data Analysis**: Comprehensive statistical analysis and insights
- **📝 Summarization**: Generate executive summaries from data
- **📊 Visualization Suggestions**: Get recommendations for charts and visualizations

### PowerPoint Generation
- **🎨 Automatic Slide Creation**: Generate presentations from Excel data
- **🎭 Multiple Themes**: Choose from 5+ professional color themes
- **✨ Smart Formatting**: Professional layouts with auto-formatted text
- **📥 Download Ready**: Export as ready-to-use .pptx files

### End-to-End Workflow
- **🔄 Complete Pipeline**: Upload Excel → Analyze → Generate PPT in one click
- **⚡ Real-time Processing**: See results instantly with Claude AI
- **🛠️ Full Customization**: Adjust titles, themes, and content

## 🚀 Quick Start

### Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/Arpan-234/App.git
   cd App
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up your API key**:
   - Get your Anthropic API key from [https://console.anthropic.com/](https://console.anthropic.com/)
   - Create `.streamlit/secrets.toml`:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-your-actual-api-key"
   ```

4. **Run the app locally**:
   ```bash
   streamlit run app.py
   ```

### Deploy to Streamlit Cloud

1. Push your repository to GitHub
2. Go to [Streamlit Cloud](https://share.streamlit.io/)
3. Click "New app" → Select your repo, branch, and `app.py`
4. In app settings, add your API key as a secret:
   - Settings → Secrets → Add `ANTHROPIC_API_KEY`
5. Deploy!

## 📖 Usage Guide

### Mode 1: Excel Automation
1. Select "📈 Excel Automation" from sidebar
2. Upload your Excel/CSV file
3. Choose a task: Clean, Analyze, Summarize, or Visualize
4. Click "Analyze" and review Claude AI's insights
5. Download the processed file

### Mode 2: PowerPoint Creation
1. Select "📊 PowerPoint Creation" from sidebar
2. Upload your Excel/CSV file
3. Enter presentation title
4. Choose a theme color
5. Click "Generate PPT"
6. Download your presentation

### Mode 3: Excel → PowerPoint (End-to-End)
1. Select "🔄 Excel → PowerPoint" from sidebar
2. Upload your file
3. Set presentation title
4. Click "Create Full Workflow"
5. Get complete analysis + presentation in one step

## 📋 Requirements

```
streamlit==1.34.0
pandas==2.1.4
openpyxl==3.11.0
python-pptx==0.6.23
anthropric==0.20.0
pydantic==2.5.3
```

## 🔧 Configuration

### Environment Variables
- `ANTHROPIC_API_KEY`: Your Claude AI API key (required)
- Optional: Set debug mode in `.streamlit/secrets.toml`

### AI Model Selection
Switch between Claude models in the sidebar:
- **claude-3-5-sonnet-20241022** (faster)
- **claude-3-opus-20250219** (more powerful)

## 📁 Project Structure

```
App/
├── app.py                 # Main Streamlit application
├── requirements.txt       # Python dependencies
├── .streamlit/
│   └── secrets.toml       # API keys (add your key here)
└── README.md              # This file
```

## 🧪 Testing Edge Cases

The app handles various edge cases:
- ✅ Empty Excel files
- ✅ Large datasets (>10k rows)
- ✅ Missing values in data
- ✅ Special characters in column names
- ✅ Various file formats (Excel, CSV)
- ✅ API timeout handling
- ✅ Invalid data type handling

## 🔒 Security

- API keys are stored securely in Streamlit Secrets
- No data is logged or stored on servers
- Files are processed locally
- Never commit secrets.toml to version control

## 📊 Example Workflow

```
1. Upload sales_data.xlsx
2. AI analyzes: 50k records, 12 columns
3. Generates insights: "20% growth in Q3"
4. Creates 5-slide presentation with:
   - Title slide
   - Executive summary
   - Key metrics
   - Trends analysis
   - Recommendations
5. Download presentation.pptx
```

## 🐛 Troubleshooting

### "ANTHROPIC_API_KEY not found"
- Ensure `.streamlit/secrets.toml` exists
- Check API key is valid at [console.anthropic.com](https://console.anthropic.com/)

### "Error processing file"
- Verify file is valid Excel/CSV
- Check file isn't corrupted
- Try smaller sample first

### Slow performance
- Reduce file size (use first 1000 rows)
- Try "claude-3-5-sonnet" for faster results
- Check internet connection

## 🎨 Customization

### Add More Themes
Edit the `colors` dictionary in app.py to add custom color schemes

### Modify AI Prompts
Edit the `call_claude()` function parameters to change AI behavior

### Add New Modes
Duplicate an `elif mode ==` block and modify the logic

## 📚 Resources

- [Streamlit Documentation](https://docs.streamlit.io/)
- [Claude API Docs](https://docs.anthropic.com/)
- [python-pptx Documentation](https://python-pptx.readthedocs.io/)
- [Pandas Documentation](https://pandas.pydata.org/docs/)

## 📄 License

MIT License - feel free to use this project!

## 👨‍💻 Author

Created with ❤️ using Streamlit & Claude AI

## 🤝 Contributing

Contributions welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to GitHub
5. Create a Pull Request

## 📞 Support

Have questions? Open an issue on GitHub!

---

**Happy automating! 🚀**
