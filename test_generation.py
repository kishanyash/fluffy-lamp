
import os
from datetime import datetime
from ppt_generator import generate_report_ppt, PPTGenerator

def test_generation():
    # Real data provided by user
    data = {
        "report_id": "test_swiggy_report_002",
        "company_name": "Swiggy Ltd",
        "nse_symbol": "SWIGGY",
        "bom_code": "544000", 
        "rating": "BUY",
        "today_date": datetime.now().strftime("%Y-%m-%d"),
        
        # Text Fields
        "management_analysis": """The management team at Swiggy Ltd is tasked with navigating a challenging landscape characterized by intense competition and rapid technological advancements in the food delivery industry. Despite the absence of detailed historical financial data, the company's current market position, as indicated by its market capitalization of ₹78,712.46 crore, suggests a significant presence in the market. However, the lack of profitability metrics such as ROE and ROCE, which stand at -255.00% and -29.20% respectively, raises concerns about the management's ability to efficiently allocate capital and generate returns.""",
        
        "industry_overview": """The Indian food delivery industry, within which Swiggy Ltd operates, is a rapidly evolving sector characterized by intense competition, dynamic consumer preferences, and significant technological advancements. This industry has seen substantial growth over the past decade, driven by increasing internet penetration, the proliferation of smartphones, and a burgeoning middle class with disposable income. As of 2023, the Indian food delivery market is estimated to be valued at approximately ₹1.5 trillion, with expectations of continued growth at a compound annual growth rate (CAGR) of around 12%.""",
        
        "market_positioning": "Strong leader in sector with competitive advantage.",
        "financial_performance": "Stable revenue and profit growth with improving margins.",
        "growth_outlook": "Future growth supported by expansion and demand.",
        "valuation_recommendation": "BUY for long term investment.",
        "key_risks": "Competition, regulation and macro risks.",
        "company_insider": "Healthy promoter and institutional holding.",
        
        "company_background": "Swiggy is India's leading on-demand convenience platform...",
        "business_model": "Hyperlocal delivery platform connecting consumers with restaurants and stores...",
        "industry_tailwinds": "Growing digital adoption and urbanization...",
        "demand_drivers": "Convenience and variety seeking behavior...",
        "industry_risks": "Regulatory changes and gig worker classification...",
        
        # Scripts
        "podcast_script": """Host A: Welcome back to "Market Masters," the show that slices through the market noise. Today, we're diving into a company that's caught a lot of eyes lately. But is it a shining star or a flickering flame about to go out?\n\nHost B: That’s right, we’re talking about TechRise, the new tech-driven consumer gadget company that’s been making headlines. I’m excited about their growth potential; they might just revolutionize how we interact with our devices.\n\nHost A: Sure, TechRise has been getting attention, but let's start with the basics. Their business model focuses on seamless integration betwe...""",
        
        "video_script": """Tata Consultancy Services, or TCS, operates in the global information technology services and consulting industry, one of the fastest-growing sectors driven by digital transformation across businesses. The IT services industry includes software development, cloud computing, cybersecurity, data analytics, artificial intelligence, and IT consulting. As companies worldwide adopt digital technologies to improve efficiency and customer experience, demand for IT services continues to grow rapidly.\n\nIndia is a global hub for IT services, with companies like TCS, Infosys, and Wipro leading the market....""",
        
        # Images
        "summary_table": "https://bmpvcjbfeyvkkbvclwkb.supabase.co/storage/v1/object/public/charts/SWIGGY/summary_chart_20260216_111325.png",
        
        "chart_custom": "https://bmpvcjbfeyvkkbvclwkb.supabase.co/storage/v1/object/public/charts/SWIGGY/custom_chart_20260216_111325.png",
        
        "price_chart": "https://bmpvcjbfeyvkkbvclwkb.supabase.co/storage/v1/object/public/charts/SWIGGY/price_chart_20260216_111325.png",
        
        # Fillers for others to avoid errors if logic expects them (though we handle None)
        "chart_profit_loss": None,
        "chart_balance_sheet": None,
        "chart_cash_flow": None,
        "chart_ratio_analysis": None
    }

    print("Starting Test Generation with Real Data...")
    
    # Ensure master_template.pptx exists
    if not os.path.exists("master_template.pptx"):
        print("Error: master_template.pptx not found!")
        return

    try:
        # We manually patch the chart custom position if needed or trust the script
        # The script now uses dynamic placement for 'price_chart' ({{prize_chart}})
        # and 'summary_table' ({{financial_table}}).
        # 'chart_custom' is hardcoded to Slide 12.
        
        output_path = generate_report_ppt(data, "master_template.pptx", "./output")
        print(f"\nSUCCESS! Report generated at: {output_path}")
        print(f"File size: {os.path.getsize(output_path)} bytes")
        
    except Exception as e:
        print(f"\nFAILED: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_generation()
