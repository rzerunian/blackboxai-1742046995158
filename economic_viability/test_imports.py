"""
Test script to verify all modules can be imported correctly
and basic functionality works without GUI.
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

def test_basic_functionality():
    print("Testing basic functionality...")
    
    # Initialize managers
    print("Initializing managers...")
    project_manager = ProjectManager()
    config = Config()
    capex_manager = CapexManager()
    opex_manager = OpexManager()
    receitas_manager = ReceitasManager()
    
    # Test config
    print("\nTesting Config...")
    success, msg = config.update(tma=10.0, ir=15.0, csll=9.0)
    print(f"Config update: {msg}")
    print(f"Effective tax rate: {config.get_effective_tax_rate()}%")
    
    # Test CapEx
    print("\nTesting CapEx...")
    success, msg = capex_manager.add_item(
        description="Equipment A",
        quantity=2,
        unit_price=1000.0,
        month=1
    )
    print(f"Adding CapEx item: {msg}")
    print(f"Total CapEx: R$ {capex_manager.total_investment:.2f}")
    
    # Test OpEx
    print("\nTesting OpEx...")
    success, msg = opex_manager.add_item(
        description="Maintenance",
        quantity=1,
        unit_price=500.0,
        recurrent=True
    )
    print(f"Adding OpEx item: {msg}")
    print(f"Total annual OpEx: R$ {opex_manager.total_annual_cost:.2f}")
    
    # Test Receitas
    print("\nTesting Receitas...")
    success, msg = receitas_manager.add_item(
        description="Service A",
        quantity=1,
        unit_price=2000.0,
        recurrent=True,
        growth_rate=2.0
    )
    print(f"Adding Receita item: {msg}")
    print(f"Total annual revenue: R$ {receitas_manager.total_annual_revenue:.2f}")
    
    # Test Financial Calculations
    print("\nTesting Financial Calculations...")
    calc = FinancialCalculations(
        capex_manager,
        opex_manager,
        receitas_manager,
        config
    )
    success, msg, results = calc.calculate_all()
    print(f"Calculations: {msg}")
    if success:
        formatted = calc.format_results()
        print("\nResults:")
        print(f"TIR: {formatted['tir']}")
        print(f"VPL: {formatted['vpl']}")
        print(f"Payback: {formatted['payback']}")
        print(f"Dívida/EBITDA: {formatted['divida_ebitda']}")

if __name__ == "__main__":
    test_basic_functionality()
