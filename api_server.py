"""
Flask API Server for PPT Generation
====================================
This server receives webhook requests from n8n and triggers
PowerPoint report generation.

Usage:
    python api_server.py

Endpoints:
    POST /generate-ppt - Generate PPT from report data
    GET /health - Health check endpoint
"""

import os
import json
from datetime import datetime
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from ppt_generator import generate_report_ppt, PPTGenerator

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

# Configuration
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(SCRIPT_DIR, "master_template.pptx")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output")

# Ensure output directory exists
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint."""
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "template_exists": os.path.exists(TEMPLATE_PATH)
    })


@app.route('/analyze-template', methods=['GET'])
def analyze_template():
    """Analyze the template and return placeholder information."""
    try:
        generator = PPTGenerator(TEMPLATE_PATH)
        generator.load_template()
        
        return jsonify({
            "success": True,
            "template_path": TEMPLATE_PATH,
            "placeholders": generator.placeholders_map,
            "slide_count": len(generator.prs.slides)
        })
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    """
    Generate a PowerPoint presentation from the provided data.
    
    Expected JSON body structure (from n8n Supabase node):
    {
        "report_id": "uuid",
        "company_name": "Company Name",
        "nse_symbol": "SYMBOL",
        "bom_code": "CODE",
        "rating": "BUY",
        "company_background": "markdown content",
        "business_model": "markdown content",
        "management_analysis": "markdown content",
        "industry_overview": "markdown content",
        "industry_tailwinds": "markdown content",
        "demand_drivers": "markdown content",
        "industry_risks": "markdown content",
        "summary_table": "url (image)",
        "chart_custom": "url (image)",
        "chart_profit_loss": "url",
        "chart_balance_sheet": "url",
        "chart_cash_flow": "url",
        "chart_ratio_analysis": "url"
    }
    """
    try:
        # Get JSON data from request
        data = request.get_json()
        
        if not data:
            return jsonify({
                "success": False,
                "error": "No JSON data provided"
            }), 400
        
        # Log incoming request
        report_id = data.get('report_id', 'unknown')
        print(f"\n{'=' * 60}")
        print(f"Received request for report: {report_id}")
        print(f"{'=' * 60}")
        
        # Debug: Log all heading and cs_ fields received
        print("\n--- DEBUG: Incoming data keys ---")
        print(f"  Total keys received: {len(data.keys())}")
        debug_fields = ['cs_masterheading', 'cs_marketing_positioning', 'cs_financial_performance',
                        'cs_grow_outlook', 'cs_value_and_recommendation', 'cs_key_risks',
                        'cs_company_insider', 'company_background_h', 'business_model_h',
                        'management_analysis_h', 'industry_overview_h', 'industry_tailwinds_h',
                        'demand_drivers_h', 'industry_risks_h']
        for field in debug_fields:
            val = data.get(field)
            if val:
                print(f"  {field}: '{str(val)[:80]}...'")
            else:
                print(f"  {field}: [EMPTY/NULL]")
        print("--- END DEBUG ---\n")
        
        # Validate required fields
        required_fields = ['report_id']
        missing_fields = [f for f in required_fields if not data.get(f)]
        
        if missing_fields:
            return jsonify({
                "success": False,
                "error": f"Missing required fields: {missing_fields}"
            }), 400
        
        # Generate the PPT
        output_path = generate_report_ppt(
            data=data,
            template_path=TEMPLATE_PATH,
            output_dir=OUTPUT_DIR
        )
        
        # Get file info
        file_name = os.path.basename(output_path)
        
        print(f"Sending file: {file_name}")
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name=file_name,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
    except FileNotFoundError as e:
        return jsonify({
            "success": False,
            "error": f"Template not found: {str(e)}"
        }), 500
        
    except Exception as e:
        print(f"Error generating PPT: {e}")
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    """Download a generated PPT file."""
    try:
        file_path = os.path.join(OUTPUT_DIR, filename)
        
        if not os.path.exists(file_path):
            return jsonify({
                "success": False,
                "error": "File not found"
            }), 404
        
        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


@app.route('/list-reports', methods=['GET'])
def list_reports():
    """List all generated reports."""
    try:
        files = []
        for filename in os.listdir(OUTPUT_DIR):
            if filename.endswith('.pptx'):
                file_path = os.path.join(OUTPUT_DIR, filename)
                files.append({
                    "filename": filename,
                    "size_bytes": os.path.getsize(file_path),
                    "created_at": datetime.fromtimestamp(
                        os.path.getctime(file_path)
                    ).isoformat()
                })
        
        return jsonify({
            "success": True,
            "count": len(files),
            "files": sorted(files, key=lambda x: x['created_at'], reverse=True)
        })
        
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


if __name__ == '__main__':
    print("=" * 60)
    print("PPT Generator API Server")
    print("=" * 60)
    print(f"\nTemplate Path: {TEMPLATE_PATH}")
    print(f"Output Directory: {OUTPUT_DIR}")
    print(f"Template Exists: {os.path.exists(TEMPLATE_PATH)}")
    print("\nStarting server on http://localhost:5000")
    print("=" * 60)
    
    # Run the Flask app
    app.run(host='0.0.0.0', port=5000, debug=True)
