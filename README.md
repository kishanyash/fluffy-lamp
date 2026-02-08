# PPT Report Generator

Automated PowerPoint presentation generator for research reports, designed to integrate with n8n workflows and Supabase data.

## 📁 Project Structure

```
report_engine/
├── master_template.pptx    # Your PowerPoint template
├── ppt_generator.py        # Core PPT generation logic
├── api_server.py           # Flask API server for webhooks
├── requirements.txt        # Python dependencies
├── output/                 # Generated reports directory
└── README.md              # This file
```

## 🚀 Quick Start

### 1. Install Dependencies

```bash
cd report_engine
pip install -r requirements.txt
```

Or install individually:
```bash
pip install python-pptx requests flask flask-cors
```

### 2. Run the Generator Directly

```bash
python ppt_generator.py
```

This will analyze your template and generate a sample PPT.

### 3. Run the API Server

```bash
python api_server.py
```

The server will start on `http://localhost:5000`

## 📡 API Endpoints

### Health Check
```
GET /health
```
Returns server status and template availability.

### Analyze Template
```
GET /analyze-template
```
Returns information about placeholders found in the template.

### Generate PPT
```
POST /generate-ppt
Content-Type: application/json

{
    "report_id": "c49b2aa1-80eb-4436-b14c-2a74d7966feb",
    "industry_overview": "# Industry Overview\n\n...",
    "industry_risks": "# Industry Risks\n\n...",
    "industry_tailwinds": "# Key Industry Tailwinds\n\n...",
    "demand_drivers": "# Demand Drivers\n\n...",
    "chart_profit_loss": "https://...",
    "chart_balance_sheet": "https://...",
    "chart_cash_flow": "https://...",
    "chart_ratio_analysis": "https://...",
    "chart_summary": "https://..."
}
```

### Download Generated File
```
GET /download/<filename>
```

### List Generated Reports
```
GET /list-reports
```

## 🔗 n8n Integration

### Step 1: Add HTTP Request Node After Supabase
In your n8n workflow, add an **HTTP Request** node after the Supabase node:

- **Method**: POST
- **URL**: `http://localhost:5000/generate-ppt` (or your server's URL)
- **Headers**: 
  - `Content-Type`: `application/json`
- **Body Type**: JSON
- **Body Content**: Use expression to pass Supabase data:
```json
{
    "report_id": "{{ $json.report_id }}",
    "industry_overview": "{{ $json.industry_overview }}",
    "industry_risks": "{{ $json.industry_risks }}",
    "industry_tailwinds": "{{ $json.industry_tailwinds }}",
    "demand_drivers": "{{ $json.demand_drivers }}",
    "chart_profit_loss": "{{ $json.chart_profit_loss }}",
    "chart_balance_sheet": "{{ $json.chart_balance_sheet }}",
    "chart_cash_flow": "{{ $json.chart_cash_flow }}",
    "chart_ratio_analysis": "{{ $json.chart_ratio_analysis }}",
    "chart_summary": "{{ $json.chart_summary }}"
}
```

### Step 2: Handle Response
The API will return:
```json
{
    "success": true,
    "message": "PPT generated successfully",
    "report_id": "c49b2aa1-...",
    "output_file": "report_c49b2aa1..._20260207_112830.pptx",
    "output_path": "C:\\...\\output\\report_....pptx",
    "file_size_bytes": 125000,
    "generated_at": "2026-02-07T11:28:30.123456"
}
```

## 🎨 Template Customization

### Using Placeholders
In your PowerPoint template, use `{{placeholder_name}}` syntax for text replacement:

- `{{company_name}}` - Company name
- `{{report_date}}` - Report date
- `{{industry_overview}}` - Industry overview content
- etc.

### Content Mapping
Edit `ppt_generator.py` to customize the `content_mapping` dictionary in the `populate_from_data` method:

```python
content_mapping = {
    'industry_overview': {
        'type': 'text',
        'placeholder': 'industry_overview',  # Placeholder name in template
        'slide_idx': 1,  # Which slide (0-indexed)
    },
    'chart_profit_loss': {
        'type': 'image',
        'slide_idx': 6,
        'position': {'left': 1.0, 'top': 1.5, 'width': 8.0}
    },
    # ... more mappings
}
```

## 📊 Supported Data Fields

| Field | Type | Description |
|-------|------|-------------|
| `report_id` | string | Unique identifier for the report |
| `industry_overview` | markdown | Industry overview content |
| `industry_risks` | markdown | Risk analysis content |
| `industry_tailwinds` | markdown | Industry growth drivers |
| `demand_drivers` | markdown | Demand analysis content |
| `chart_summary` | URL | Summary chart image |
| `chart_profit_loss` | URL | P&L chart image |
| `chart_balance_sheet` | URL | Balance sheet chart |
| `chart_cash_flow` | URL | Cash flow chart |
| `chart_ratio_analysis` | URL | Financial ratios chart |

## 🐛 Troubleshooting

### Template Not Found
Ensure `master_template.pptx` is in the same directory as the Python scripts.

### Image Download Fails
- Check if URLs are accessible
- Verify network connectivity
- Check for firewall restrictions

### Placeholder Not Replaced
- Run `/analyze-template` to see detected placeholders
- Ensure placeholder syntax is `{{name}}` (double braces)
- Check for typos in placeholder names

## 📝 License

MIT License - Feel free to use and modify.
