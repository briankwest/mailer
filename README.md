# Letter Generation System for #10 Windowed Envelopes

A professional letter generation system that creates perfectly formatted PDF letters for standard #10 double-window envelopes, complete with fold lines, multi-page text flow, and Mac-native printing support.

## Features

- **Precise Window Alignment**: Addresses positioned exactly for #10 envelope windows
- **Fold Line Markers**: 4mm guide lines in margins at 3.67" and 7.33" for perfect tri-fold
- **Multi-Page Support**: Automatic text flow with headers/footers on all pages
- **Professional Formatting**: Multiple font options (Times New Roman, Helvetica, Courier)
- **Mac Printing Integration**: Direct printing, Preview.app integration, and print dialog support
- **Signature Support**: Embed PNG signatures or use typed signatures
- **JSON Configuration**: Complete control over letter content and formatting
- **Letter Types**: Business, formal, congressional, and personal letter templates

## Quick Start

### 1. Setup

```bash
# Clone or download the project
cd mailer

# Create virtual environment
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 2. Generate Your First Letter

```bash
# Generate a business letter
python mailer.py examples/sample_letters/business_letter.json

# Generate a congressional letter
python mailer.py examples/sample_letters/congressional_letter.json
```

The PDF will be generated and automatically opened in Preview.app.

## Usage

### Basic Commands

```bash
# Generate and preview a letter (default)
python mailer.py letter.json

# Save to specific location
python mailer.py letter.json -o my_letter.pdf

# Validate JSON without generating
python mailer.py letter.json --validate

# List available printers
python mailer.py --list-printers

# Send directly to printer
python mailer.py letter.json --print

# Open print dialog
python mailer.py letter.json --print-dialog

# Print to specific printer
python mailer.py letter.json --print --printer "Brother_MFC_L3750CDW"
```

### Command Line Options

| Option | Description |
|--------|-------------|
| `-o, --output` | Specify output PDF path |
| `--preview` | Open in Preview.app (default) |
| `--print` | Send directly to printer |
| `--print-dialog` | Open Mac print dialog |
| `--printer NAME` | Specify printer name |
| `--list-printers` | List available printers |
| `--validate` | Validate JSON only |
| `--font NAME` | Override font selection |

## JSON Configuration

### Basic Structure

```json
{
  "metadata": {
    "type": "business",
    "date": "2025-10-24",
    "date_format": "full",
    "reference_id": "LETTER_001"
  },
  "positioning": {
    "margins": {...},
    "return_address": {...},
    "recipient_address": {...}
  },
  "return_address": {
    "name": "John Doe",
    "organization": "ACME Corp",
    "street_1": "123 Main St",
    "city": "New York",
    "state": "NY",
    "zip": "10001"
  },
  "recipient_address": {
    "name": "Jane Smith",
    "organization": "TechCorp",
    "street_1": "456 Innovation Dr",
    "city": "San Francisco",
    "state": "CA",
    "zip": "94105"
  },
  "content": {
    "salutation": "Dear Ms. Smith",
    "body": [
      "First paragraph...",
      "Second paragraph..."
    ],
    "closing": "Sincerely",
    "signature": {...}
  },
  "formatting": {
    "font_family": "Times-Roman",
    "font_size": 12,
    "line_spacing": 1.5
  },
  "fold_lines": {
    "enabled": true,
    "positions": [3.67, 7.33]
  },
  "header": {...},
  "footer": {...}
}
```

### Key Configuration Options

#### Letter Types
- `business` - Standard business letter
- `formal` - Formal personal letter
- `congressional` - Letters to representatives/senators
- `personal` - Casual personal letter

#### Date Formatting
- `full` - "October 24, 2025"
- `abbreviated` - "Oct 24, 2025"
- `custom` - Custom format string

#### Available Fonts
- `Times-Roman` - Times New Roman (default)
- `Helvetica` - Helvetica
- `Courier` - Courier

#### Fold Lines
The system automatically adds 4mm guide lines in the margins at:
- **3.67 inches** from top (bottom fold)
- **7.33 inches** from top (top fold)

These create perfect tri-folds for #10 envelopes.

## Envelope Specifications

### #10 Double Window Envelope
- **Envelope Size**: 9.5" × 4.125"
- **Top Window** (Return Address): 0.5" from left, 0.625" from top, 3.5" × 1.0"
- **Bottom Window** (Recipient): 0.75" from left, 2.0625" from top, 4.0" × 1.125"

## Page Layout

### First Page
- Return address (positioned for top window)
- Date (formatted as specified)
- Recipient address (positioned for bottom window)
- Subject line (optional, bold)
- Salutation
- Body text begins

### Subsequent Pages
- Header: Recipient name, page number, date
- Full page of body text
- Footer: Page X of Y

### Signature Options

#### PNG Image Signature
```json
"signature": {
  "type": "image",
  "image_path": "assets/signatures/signature.png",
  "width": 2.0,
  "height": 0.75,
  "typed_name": "John Doe",
  "title": "Director"
}
```

#### Typed Signature
```json
"signature": {
  "type": "typed",
  "typed_name": "John Doe",
  "title": "Director"
}
```

## Advanced Features

### Multi-Page Text Flow
The system automatically flows text across multiple pages with:
- Widow/orphan control
- Proper paragraph spacing
- Headers on continuation pages
- Consistent margins

### Congressional Letters
Special formatting for letters to representatives:
```json
"recipient_address": {
  "honorific": "The Honorable",
  "name": "Charles E. Schumer",
  "title": "United States Senator",
  "organization": "United States Senate",
  "street_1": "322 Hart Senate Office Building",
  "city": "Washington",
  "state": "DC",
  "zip": "20510"
}
```

### Enclosures and CC
```json
"content": {
  ...
  "postscript": "P.S. Additional information enclosed.",
  "enclosures": [
    "Annual Report 2025",
    "Product Brochure"
  ],
  "cc": [
    "John Smith, CFO",
    "Jane Doe, Legal Counsel"
  ]
}
```

## Testing Your Letters

### 1. Visual Verification
- Generate PDF and open in Preview
- Check address positioning in windows
- Verify fold lines appear in margins
- Confirm headers/footers on all pages

### 2. Physical Test
1. Print the letter
2. Fold at the marked lines (4mm guides in margins)
3. Insert into #10 envelope
4. Verify addresses appear in windows

### 3. JSON Validation
```bash
python mailer.py your_letter.json --validate
```

## Troubleshooting

### Common Issues

**Q: Addresses don't align with windows**
- Check positioning values in JSON
- Ensure using correct envelope type (#10 double window)
- Verify printer scaling is set to 100%

**Q: Fold lines not visible**
- Ensure `fold_lines.enabled` is `true`
- Check printer doesn't crop margins
- Lines are intentionally subtle (light gray)

**Q: Font not found**
- Use one of the standard fonts (Times-Roman, Helvetica, etc.)
- Check font name spelling in JSON

**Q: Multi-page text cutoff**
- System automatically handles page breaks
- Check margins and spacing settings
- Ensure content array is properly formatted

## Project Structure

```
mailer/
├── mailer.py              # Main program (unified)
├── requirements.txt       # Python dependencies
├── PLAN.md               # Development plan
├── README.md             # This file
├── config/
│   ├── schema.json       # JSON schema definition
│   ├── fonts.json        # Font configurations
│   └── templates/        # Letter templates
├── assets/
│   ├── signatures/       # PNG signature files
│   └── logos/           # Optional logos
├── examples/
│   └── sample_letters/   # Example JSON files
├── output/              # Generated PDFs
└── tests/               # Test suite
```

## Requirements

- **macOS** (for printing features)
- **Python 3.8+**
- **Dependencies**: reportlab, pillow, pydantic, click

## License

MIT License

Copyright (c) 2025 Brian West

See [LICENSE](LICENSE) file for full details.

## Contributing

Contributions welcome! Please follow the existing code style and add tests for new features.

## Support

For issues or questions, please create an issue on the project repository.

---

Built with ❤️ for professional correspondence