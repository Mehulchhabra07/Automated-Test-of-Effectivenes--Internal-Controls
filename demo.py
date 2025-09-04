#!/usr/bin/env python3
"""
🚀 Demo Script for AI-Powered TOE Evidence Analysis Tool

This interactive demonstration showcases the capabilities of the AI-driven
evidence analysis framework. Perfect for understanding the tool's
functionality before analyzing your own evidence data.

Author: Mehul Chhabra
GitHub: https://github.com/Mehulchhabra07/AI-Powered-TOE-Evidence-Analysis
"""

import os
import sys
from pathlib import Path

def print_banner():
    """Display an attractive banner for the demo"""
    print("🚀" + "="*78 + "🚀")
    print("    AI-POWERED TOE EVIDENCE ANALYSIS TOOL - INTERACTIVE DEMO")
    print("🚀" + "="*78 + "🚀")
    print()
    print("📋 This demo will:")
    print("   • Check your environment setup")
    print("   • Create sample control data")
    print("   • Show evidence processing capabilities")
    print("   • Run a complete analysis")
    print()

def setup_demo():
    """Set up the demo environment with comprehensive checks"""
    print_banner()
    print("🔧 Performing environment validation...\n")
    
    # Check Python version
    python_version = sys.version_info
    if python_version < (3, 8):
        print("❌ Python 3.8+ required. Current version:", 
              f"{python_version.major}.{python_version.minor}")
        return False
    print(f"✅ Python {python_version.major}.{python_version.minor} detected")
    
    # Check dependencies
    required_packages = ['pandas', 'openpyxl', 'openai', 'httpx', 'python-docx', 'PyPDF2']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package.replace('-', '_'))
            print(f"✅ {package} installed")
        except ImportError:
            missing_packages.append(package)
            print(f"❌ {package} missing")
    
    if missing_packages:
        print(f"\n⚠️  Missing packages: {', '.join(missing_packages)}")
        print("Please install them with: pip install -r requirements.txt")
        return False
    
    # Check OpenAI API key
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        print("\n⚠️  OpenAI API key not found!")
        print("   Set your API key with:")
        print("   export OPENAI_API_KEY='your-api-key-here'")
        print("\n   Or for this session only:")
        api_key = input("   Enter your OpenAI API key (or press Enter to skip): ").strip()
        
        if api_key:
            os.environ["OPENAI_API_KEY"] = api_key
            print("   ✅ API key set for this session")
        else:
            print("   ⚠️  Continuing without API key (analysis will fail)")
            return False
    else:
        masked_key = f"{api_key[:7]}...{api_key[-4:]}" if len(api_key) > 11 else "***"
        print(f"✅ OpenAI API key found: {masked_key}")
    
    # Check if sample file exists
    sample_file = Path("sample_controls.xlsx")
    if not sample_file.exists():
        print("\n📄 Sample file not found, creating one...")
        if create_sample_file():
            print("   ✅ Sample file created successfully")
        else:
            print("   ❌ Failed to create sample file")
            return False
    else:
        print("✅ Sample file 'sample_controls.xlsx' found")
    
    # Create evidence folder structure
    evidence_dir = Path("Evidence")
    if not evidence_dir.exists():
        print("\n📂 Creating evidence folder structure...")
        if create_evidence_structure():
            print("   ✅ Evidence folders created successfully")
        else:
            print("   ❌ Failed to create evidence structure")
            return False
    else:
        print("✅ Evidence folder structure exists")
    
    print("\n🎉 Demo environment ready!")
    print("=" * 60)
    return True

def create_sample_file():
    """Create a comprehensive sample Excel file for demonstration"""
    try:
        import pandas as pd
        
        # Sample control data focused on evidence analysis
        data = {
            'Risk': ['R001', 'R002', 'R003', 'R004'],
            'Risk Description': [
                'Risk of unauthorized access to financial systems resulting in data manipulation or theft',
                'Risk of incomplete or inaccurate financial reporting due to manual process errors',
                'Risk of excessive spending without proper authorization and budget oversight',
                'Risk of data loss or corruption affecting business continuity and compliance'
            ],
            'Control': ['C001', 'C002', 'C003', 'C004'],
            'Control Description': [
                'Monthly access review of all financial system users with formal documentation and manager approval for any changes',
                'Automated system validation checks are performed in real-time on all financial entries with exception reporting to the Finance Manager',
                'Department heads review and approve all expenses above $1,000 using digital approval workflow with documented business justification',
                'IT team performs weekly automated backups of critical financial data with monthly restore testing and documented procedures'
            ]
        }
        
        df = pd.DataFrame(data)
        df.to_excel('sample_controls.xlsx', index=False, engine='openpyxl')
        
        print(f"   📊 Created {len(df)} sample controls")
        return True
        
    except Exception as e:
        print(f"   ❌ Error creating sample file: {e}")
        return False

def create_evidence_structure():
    """Create sample evidence folders with demonstration files"""
    try:
        evidence_dir = Path("Evidence")
        evidence_dir.mkdir(exist_ok=True)
        
        # Create evidence folders for each control
        controls = ["C001", "C002", "C003", "C004"]
        
        for control in controls:
            control_dir = evidence_dir / control
            control_dir.mkdir(exist_ok=True)
            
            # Create sample evidence files
            sample_files = {
                f"{control}_access_review.txt": f"Monthly Access Review Report for {control}\n\nDate: 2024-03-15\nReviewer: Jane Smith\nSystem: Financial ERP\n\nReview Summary:\n- Total users reviewed: 45\n- Access changes: 3\n- Approvals obtained: Yes\n- Documentation complete: Yes\n\nConclusion: Control operating effectively",
                f"{control}_evidence.txt": f"Supporting Evidence for {control}\n\nThis file contains evidence of control execution including:\n- Process documentation\n- Approval workflows\n- Test results\n- Compliance verification\n\nAll requirements have been met according to policy.",
                "README.txt": f"Evidence folder for {control}\n\nThis folder contains evidence files demonstrating the effective operation of the control.\nFiles are organized by date and evidence type."
            }
            
            for filename, content in sample_files.items():
                file_path = control_dir / filename
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
        
        # Create an additional folder with a .gitkeep for empty state
        (evidence_dir / ".gitkeep").touch()
        
        print(f"   📂 Created evidence folders for {len(controls)} controls")
        return True
        
    except Exception as e:
        print(f"   ❌ Error creating evidence structure: {e}")
        return False

def display_analysis_preview():
    """Show what the analysis will evaluate"""
    print("\n🔍 ANALYSIS PREVIEW")
    print("=" * 60)
    print("The AI will analyze evidence across multiple dimensions:")
    print()
    
    analysis_areas = [
        ("📄 Evidence Summary", "Comprehensive overview of all evidence found"),
        ("🔍 Document Analysis", "Key information extraction from all file types"),
        ("✅ Effectiveness Assessment", "Professional evaluation of control operation"),
        ("🎯 Sufficiency Analysis", "Whether evidence proves control effectiveness"),
        ("📋 Gap Identification", "Missing elements or improvement areas"),
        ("🏆 Professional Conclusion", "Audit-quality YES/NO determination"),
        ("📊 Detailed Reasoning", "Comprehensive auditor-style explanations")
    ]
    
    for area, description in analysis_areas:
        print(f"   {area}: {description}")
    
    print("\n📈 EXPECTED OUTPUT")
    print("=" * 60)
    print("📄 Excel report with:")
    print("   • Detailed evidence summary for each control")
    print("   • Professional sufficiency assessment")
    print("   • Color-coded results and insights")
    print("   • Audit-trail quality documentation")
    print("   • Professional formatting and styling")

def run_demo():
    """Run the complete demonstration"""
    if not setup_demo():
        print("\n❌ Demo setup failed. Please resolve the issues above.")
        return False
    
    display_analysis_preview()
    
    print("\n🚀 STARTING ANALYSIS")
    print("=" * 60)
    
    # Get user confirmation
    while True:
        response = input("\nProceed with AI evidence analysis? (y/n): ").lower().strip()
        if response in ['y', 'yes']:
            break
        elif response in ['n', 'no']:
            print("Demo cancelled by user.")
            return True
        else:
            print("Please enter 'y' for yes or 'n' for no.")
    
    print("\n🔄 Running evidence analysis...")
    print("   This may take 3-7 minutes depending on API response times...")
    print("   Please be patient while the AI analyzes each control's evidence...")
    print()
    
    try:
        # Import and run the main analysis
        from toe_evidence_analysis_enhanced import main
        main()
        
        print("\n" + "=" * 60)
        print("🎉 DEMO COMPLETED SUCCESSFULLY!")
        print("=" * 60)
        print()
        print("📊 Check the output file: sample_controls_TOE_EvidenceAnalysis.xlsx")
        print("📂 Review the Evidence/ folder structure")
        print("📋 Examine the detailed AI analysis results")
        print()
        print("🚀 Ready to analyze your own evidence data!")
        print("   1. Replace sample_controls.xlsx with your data")
        print("   2. Add your evidence files to Evidence/ folders")
        print("   3. Run: python toe_evidence_analysis_enhanced.py")
        print()
        return True
        
    except Exception as e:
        print(f"\n❌ Demo analysis failed: {e}")
        print("\nThis might be due to:")
        print("   • Missing OpenAI API key")
        print("   • Network connectivity issues")
        print("   • Invalid API key")
        print("   • API rate limits")
        return False

def main():
    """Main demo function"""
    print("Starting AI-Powered TOE Evidence Analysis Demo...\n")
    
    try:
        success = run_demo()
        if success:
            print("\n✅ Demo completed successfully!")
        else:
            print("\n⚠️  Demo completed with issues.")
    except KeyboardInterrupt:
        print("\n\n⚠️  Demo interrupted by user.")
    except Exception as e:
        print(f"\n❌ Demo failed: {e}")

if __name__ == "__main__":
    main()
