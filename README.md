# PPT Translator (Formatting Intact) with OCR Support 🎯

A powerful PowerPoint translation tool that preserves all formatting while translating content using AI. This tool maintains fonts, colors, layouts, and other styling elements while providing accurate translations between languages. **NEW**: Now includes OCR support to extract and translate text from images!

## ✨ Features

- **OCR Image Text Translation**: Extract text from images and translate it back onto the images
- Preserves all PowerPoint formatting during translation
- Supports tables, text boxes, images, and other PowerPoint elements
- Maintains font styles, sizes, colors, and alignments
- Intelligent text chunking for better translation quality
- Caches translations to avoid duplicate API calls
- Multi-threaded processing for faster execution
- Creates intermediate backups during translation
- Supports custom source and target languages
- Advanced image text detection with confidence scoring

## 🚀 Installation

1. Clone this repository:
```bash
git clone https://github.com/tristan-mcinnis/PPT-Translator-Formatting-Intact-with-Deepseek.git
cd PPT-Translator-Formatting-Intact-with-Deepseek
```

2. Create and activate a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install required packages:
```bash
pip install -r requirements.txt
```

4. **Install OCR Dependencies** (for image text translation):
```bash
# Run the setup script to install OCR dependencies
python setup_ocr.py

# Or manually install Tesseract OCR:
# Ubuntu/Debian: sudo apt-get install tesseract-ocr tesseract-ocr-eng
# macOS: brew install tesseract
# Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki
```

5. Create a `.env` file in the project root and add your API key:
```
OPENAI_API_KEY=your_openai_api_key_here
```

## 💻 Usage

1. Run the script:
```bash
python main.py
```

2. Follow the prompts:
   - Enter the path to your PowerPoint file
   - Specify source language code (default: 'zh' for Chinese)
   - Specify target language code (default: 'en' for English)
   - Choose whether to enable OCR for image text translation

3. The script will:
   - Generate an XML representation of your PowerPoint
   - Extract text from images using OCR (if enabled)
   - Translate all text content while preserving formatting
   - Overlay translated text back onto images
   - Create a new PowerPoint file with the translated content
   - Save the output file as `{original_filename}_translated.pptx`

## 📝 Example

```bash
=== PowerPoint Translator with OCR Support ===
OCR Available: Yes

Enter path to a PPTX file OR directory: /path/to/your/presentation.pptx
Enter source language code (default 'zh'): zh
Enter target language code (default 'en'): en
Enable OCR for images? (y/n, default 'y'): y

OCR Settings:
- Text confidence threshold: 30%
- Will extract text from images and translate it
- Translated text will be overlaid back onto images
```

## ⚙️ Supported Languages

The tool supports all languages available through the OpenAI API. Common language codes include:
- 'zh': Chinese
- 'en': English
- 'es': Spanish
- 'fr': French
- 'de': German
- 'ja': Japanese
- 'ko': Korean

## 🖼️ OCR Requirements

For image text extraction and translation, you need:

### Required Software:
- **Tesseract OCR**: System-level OCR engine
- **Python packages**: pytesseract, Pillow, opencv-python, numpy

### Installation:
```bash
# Automated setup
python setup_ocr.py

# Manual installation:
# Ubuntu/Debian
sudo apt-get install tesseract-ocr tesseract-ocr-eng tesseract-ocr-chi-sim

# macOS
brew install tesseract

# Windows
# Download from: https://github.com/UB-Mannheim/tesseract/wiki
```

### Troubleshooting OCR:
- **"Tesseract not found"**: Add Tesseract to your system PATH
- **Poor OCR accuracy**: Ensure images have good contrast and readable text
- **Missing language packs**: Install additional Tesseract language packages for your source language
- **Image replacement issues**: Some complex PowerPoint structures may not support automatic image replacement

## 🔍 Notes

### OCR Features:
- **Image Text Detection**: Automatically detects and translates text in images
- **Smart Image Processing**: Handles multiple image formats (PNG, JPEG, etc.) and skips unsupported formats (WMF, EMF)
- **Image Text Overlay**: Translated text is overlaid back onto images with automatic font sizing
- **Background Detection**: Smart background color detection for better text visibility
- **Confidence Filtering**: Only processes OCR text with >30% confidence for accuracy

### Formatting Preservation:
- **Enhanced Color Detection**: Supports RGB, theme colors, and brightness variants with intelligent fallbacks
- **Intelligent Text Fitting**: Automatically adjusts font sizes based on text length changes
  - Longer translations: Font size reduced proportionally (minimum 8pt)
  - Shorter translations: Font size slightly increased (maximum 150% of original)
- **Complete Property Capture**: Preserves all text formatting including fonts, colors, alignment, spacing
- **Table Cell Optimization**: Special handling for table cells with width-aware font sizing

### Technical Details:
- Arial font is used by default for better cross-platform compatibility
- Intermediate XML files are automatically cleaned up after successful processing
- OCR requires Tesseract to be installed on your system
- Rate limiting protection with exponential backoff for API stability
- Thread-safe translation caching to improve performance and reduce API costs

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! Feel free to check [issues page](https://github.com/tristan-mcinnis/PPT-Translator-Formatting-Intact-with-Deepseek/issues).

## ⭐️ Show your support

Give a ⭐️ if this project helped you!

# PPTX-Translator
