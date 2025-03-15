"""
Test script to verify Excel export functionality works correctly.
"""

import os
import sys

# Add the parent directory to Python path to allow imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from economic_viability.project_manager import ProjectManager
from economic_viability.config import Config
from economic_viability.financial_calculations import FinancialCalculations
from economic_viability.excel_exporter import ExcelExporter
from economic_viability.financial_items import (
    CapexManager,
    OpexManager,
    ReceitasManager
)

def test_excel_export():
    print("Testing Excel export functionality...")
    
    # Initialize managers
    print("Initializing managers...")
    project_manager = ProjectManager()
    config = Config()
    capex_manager = CapexManager()
    opex_manager = OpexManager()
    receitas_manager = ReceitasManager()
    
    # Configure test data
    config.update(tma=10.0, ir=15.0, csll=9.0)
    
    # Add test items
    capex_manager.add_item(
        description="Equipment A",
        quantity=2,
        unit_price=1000.0,
        month=1
    )
    
    opex_manager.add_item(
        description="Maintenance",
        quantity=1,
        unit_price=500.0,
        recurrent=True
    )
    
    receitas_manager.add_item(
        description="Service A",
        quantity=1,
        unit_price=2000.0,
        recurrent=True,
        growth_rate=2.0
    )
    
    # Calculate financials
    calc = FinancialCalculations(
        capex_manager,
        opex_manager,
        receitas_manager,
        config
    )
    calc.calculate_all()
    
    # Create Excel exporter
    exporter = ExcelExporter(
        project_manager,
        capex_manager,
        opex_manager,
        receitas_manager,
        config,
        calc
    )
    
    # Test export
    export_path = "test_export.xlsx"
    success, message = exporter.export_project(export_path)
    print(f"\nExport result: {message}")
    
    if success and os.path.exists(export_path):
        print(f"Excel file created successfully at: {export_path}")
        print(f"File size: {os.path.getsize(export_path):,} bytes")
    else:
        print("Failed to create Excel file")

if __name__ == "__main__":
    test_excel_export()
