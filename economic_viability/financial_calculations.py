import numpy as np
from datetime import datetime

class FinancialCalculations:
    def __init__(self, capex_manager, opex_manager, receitas_manager, config):
        """
        Initialize Financial Calculations with required managers.
        
        Args:
            capex_manager: Instance of CapexManager
            opex_manager: Instance of OpexManager
            receitas_manager: Instance of ReceitasManager
            config: Instance of Config
        """
        self.capex_manager = capex_manager
        self.opex_manager = opex_manager
        self.receitas_manager = receitas_manager
        self.config = config
        self.results = {}

    def calculate_all(self):
        """
        Calculate all financial indicators.
        
        Returns:
            dict: Dictionary containing all calculated indicators
        """
        try:
            # Calculate cash flows
            cash_flows = self.calculate_cash_flows()
            
            # Calculate indicators
            self.results = {
                "cash_flows": cash_flows,
                "tir": self.calculate_tir(cash_flows),
                "vpl": self.calculate_vpl(cash_flows),
                "payback": self.calculate_payback(cash_flows),
                "divida_ebitda": self.calculate_divida_ebitda(cash_flows),
                "calculated_at": datetime.now().isoformat()
            }
            
            return True, "Cálculos realizados com sucesso", self.results
            
        except Exception as e:
            return False, f"Erro ao realizar cálculos: {str(e)}", None

    def calculate_cash_flows(self):
        """
        Calculate monthly cash flows.
        
        Returns:
            list: List of monthly cash flows
        """
        cash_flows = []
        
        for month in range(1, 13):
            # Calculate revenues
            revenue = self.receitas_manager.get_monthly_revenue(month)
            
            # Calculate operational costs
            opex = self.opex_manager.get_monthly_cost(month)
            
            # Calculate capital expenditure
            capex = self.capex_manager.get_monthly_investment(month)
            
            # Calculate EBITDA (Earnings Before Interest, Taxes, Depreciation, and Amortization)
            ebitda = revenue - opex
            
            # Calculate taxes
            tax_rate = self.config.get_effective_tax_rate() / 100
            taxes = max(0, ebitda * tax_rate)
            
            # Calculate net cash flow
            net_cash_flow = ebitda - taxes - capex
            
            cash_flows.append({
                "month": month,
                "revenue": revenue,
                "opex": opex,
                "capex": capex,
                "ebitda": ebitda,
                "taxes": taxes,
                "net_cash_flow": net_cash_flow
            })
            
        return cash_flows

    def calculate_tir(self, cash_flows):
        """
        Calculate TIR (Taxa Interna de Retorno / Internal Rate of Return).
        
        Args:
            cash_flows (list): List of monthly cash flows
            
        Returns:
            float: TIR value as percentage
        """
        try:
            # Extract net cash flows
            flows = [cf["net_cash_flow"] for cf in cash_flows]
            
            # Calculate IRR using numpy
            irr = np.irr(flows)
            
            # Convert to annual rate and percentage
            annual_irr = ((1 + irr) ** 12 - 1) * 100
            
            return annual_irr
            
        except Exception:
            return None

    def calculate_vpl(self, cash_flows):
        """
        Calculate VPL (Valor Presente Líquido / Net Present Value).
        
        Args:
            cash_flows (list): List of monthly cash flows
            
        Returns:
            float: VPL value
        """
        try:
            # Get monthly discount rate from TMA
            monthly_rate = (1 + self.config.tma/100) ** (1/12) - 1
            
            # Calculate NPV
            npv = 0
            for i, cf in enumerate(cash_flows):
                npv += cf["net_cash_flow"] / (1 + monthly_rate) ** (i + 1)
                
            return npv
            
        except Exception:
            return None

    def calculate_payback(self, cash_flows):
        """
        Calculate Payback period.
        
        Args:
            cash_flows (list): List of monthly cash flows
            
        Returns:
            float: Payback period in months
        """
        try:
            cumulative_flow = 0
            for i, cf in enumerate(cash_flows):
                cumulative_flow += cf["net_cash_flow"]
                if cumulative_flow >= 0:
                    # Interpolate for more precise payback period
                    if i > 0:
                        prev_cf = cumulative_flow - cf["net_cash_flow"]
                        fraction = -prev_cf / cf["net_cash_flow"]
                        return i + fraction
                    return i + 1
            return None  # Payback not achieved within the period
            
        except Exception:
            return None

    def calculate_divida_ebitda(self, cash_flows):
        """
        Calculate Dívida Líquida sobre EBITDA.
        
        Args:
            cash_flows (list): List of monthly cash flows
            
        Returns:
            float: Dívida Líquida/EBITDA ratio
        """
        try:
            # Calculate total EBITDA
            total_ebitda = sum(cf["ebitda"] for cf in cash_flows)
            if total_ebitda == 0:
                return None
            
            # Calculate net debt (assuming it's the negative of total capex)
            total_capex = sum(cf["capex"] for cf in cash_flows)
            
            # Calculate ratio
            return total_capex / total_ebitda if total_ebitda != 0 else None
            
        except Exception:
            return None

    def format_results(self):
        """
        Format calculation results for display.
        
        Returns:
            dict: Formatted results with proper labels and units
        """
        if not self.results:
            return {
                "tir": "N/A",
                "vpl": "N/A",
                "payback": "N/A",
                "divida_ebitda": "N/A"
            }

        # Format TIR
        tir = self.results.get("tir")
        tir_formatted = f"{tir:.2f}%" if tir is not None else "N/A"

        # Format VPL
        vpl = self.results.get("vpl")
        vpl_formatted = f"R$ {vpl:,.2f}" if vpl is not None else "N/A"

        # Format Payback
        payback = self.results.get("payback")
        if payback is not None:
            years = int(payback // 12)
            months = int(payback % 12)
            payback_formatted = f"{years} anos e {months} meses"
        else:
            payback_formatted = "N/A"

        # Format Dívida/EBITDA
        divida_ebitda = self.results.get("divida_ebitda")
        divida_ebitda_formatted = f"{divida_ebitda:.2f}x" if divida_ebitda is not None else "N/A"

        return {
            "tir": tir_formatted,
            "vpl": vpl_formatted,
            "payback": payback_formatted,
            "divida_ebitda": divida_ebitda_formatted
        }

    def export_to_excel(self, workbook):
        """
        Export calculation results to Excel.
        
        Args:
            workbook: openpyxl workbook object
            
        Returns:
            tuple: (bool, str) - (success, message)
        """
        try:
            # Create Results sheet
            if "Resultados" in workbook.sheetnames:
                ws = workbook["Resultados"]
            else:
                ws = workbook.create_sheet("Resultados")

            # Clear existing content
            for row in ws.rows:
                for cell in row:
                    cell.value = None

            # Write indicators
            ws.append(["Indicador", "Valor"])
            formatted_results = self.format_results()
            ws.append(["TIR", formatted_results["tir"]])
            ws.append(["VPL", formatted_results["vpl"]])
            ws.append(["Payback", formatted_results["payback"]])
            ws.append(["Dívida/EBITDA", formatted_results["divida_ebitda"]])

            # Add cash flows table if available
            if "cash_flows" in self.results:
                ws.append([])  # Empty row
                ws.append(["Fluxo de Caixa Mensal"])
                headers = ["Mês", "Receitas", "OpEx", "CapEx", "EBITDA", "Impostos", "Fluxo Líquido"]
                ws.append(headers)

                for cf in self.results["cash_flows"]:
                    ws.append([
                        cf["month"],
                        cf["revenue"],
                        cf["opex"],
                        cf["capex"],
                        cf["ebitda"],
                        cf["taxes"],
                        cf["net_cash_flow"]
                    ])

            return True, "Resultados exportados com sucesso"
            
        except Exception as e:
            return False, f"Erro ao exportar resultados: {str(e)}"
