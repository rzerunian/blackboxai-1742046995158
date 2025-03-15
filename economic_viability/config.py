import json
import os
from datetime import datetime

class Config:
    def __init__(self):
        """Initialize configuration with default values."""
        self.tma = 0.0  # Taxa Mínima de Atratividade (%)
        self.ir = 0.0   # Imposto de Renda (%)
        self.csll = 0.0 # Contribuição Social sobre Lucro Líquido (%)
        self.modified_at = datetime.now()

    def update(self, tma=None, ir=None, csll=None):
        """
        Update configuration values.
        
        Args:
            tma (float, optional): New TMA value
            ir (float, optional): New IR value
            csll (float, optional): New CSLL value
            
        Returns:
            tuple: (bool, str) - (success, message)
        """
        try:
            if tma is not None:
                tma_value = float(tma)
                if tma_value < 0 or tma_value > 100:
                    return False, "TMA deve estar entre 0 e 100%"
                self.tma = tma_value

            if ir is not None:
                ir_value = float(ir)
                if ir_value < 0 or ir_value > 100:
                    return False, "IR deve estar entre 0 e 100%"
                self.ir = ir_value

            if csll is not None:
                csll_value = float(csll)
                if csll_value < 0 or csll_value > 100:
                    return False, "CSLL deve estar entre 0 e 100%"
                self.csll = csll_value

            self.modified_at = datetime.now()
            return True, "Configurações atualizadas com sucesso"

        except ValueError:
            return False, "Valores inválidos. Use apenas números"
        except Exception as e:
            return False, f"Erro ao atualizar configurações: {str(e)}"

    def get_effective_tax_rate(self):
        """
        Calculate the effective combined tax rate.
        
        Returns:
            float: Combined tax rate as a percentage
        """
        return self.ir + self.csll

    def to_dict(self):
        """
        Convert configuration to dictionary.
        
        Returns:
            dict: Dictionary representation of the configuration
        """
        return {
            "tma": self.tma,
            "ir": self.ir,
            "csll": self.csll,
            "modified_at": self.modified_at.isoformat()
        }

    def from_dict(self, data):
        """
        Load configuration from dictionary.
        
        Args:
            data (dict): Dictionary containing configuration data
            
        Returns:
            tuple: (bool, str) - (success, message)
        """
        try:
            success, message = self.update(
                tma=data.get("tma"),
                ir=data.get("ir"),
                csll=data.get("csll")
            )
            
            if success and "modified_at" in data:
                self.modified_at = datetime.fromisoformat(data["modified_at"])
                
            return success, message
            
        except Exception as e:
            return False, f"Erro ao carregar configurações: {str(e)}"

    def save_to_file(self, filepath):
        """
        Save configuration to a JSON file.
        
        Args:
            filepath (str): Path to save the configuration file
            
        Returns:
            tuple: (bool, str) - (success, message)
        """
        try:
            with open(filepath, 'w') as f:
                json.dump(self.to_dict(), f, indent=4)
            return True, "Configurações salvas com sucesso"
        except Exception as e:
            return False, f"Erro ao salvar configurações: {str(e)}"

    def load_from_file(self, filepath):
        """
        Load configuration from a JSON file.
        
        Args:
            filepath (str): Path to the configuration file
            
        Returns:
            tuple: (bool, str) - (success, message)
        """
        try:
            if not os.path.exists(filepath):
                return False, "Arquivo de configuração não encontrado"
                
            with open(filepath, 'r') as f:
                data = json.load(f)
            
            return self.from_dict(data)
            
        except json.JSONDecodeError:
            return False, "Arquivo de configuração inválido"
        except Exception as e:
            return False, f"Erro ao carregar configurações: {str(e)}"

    def save_to_excel(self, workbook, sheet_name="Configuração"):
        """
        Save configuration to an Excel worksheet.
        
        Args:
            workbook: openpyxl workbook object
            sheet_name (str): Name of the worksheet
            
        Returns:
            tuple: (bool, str) - (success, message)
        """
        try:
            if sheet_name in workbook.sheetnames:
                ws = workbook[sheet_name]
            else:
                ws = workbook.create_sheet(sheet_name)

            # Clear existing content
            for row in ws.rows:
                for cell in row:
                    cell.value = None

            # Write headers and values
            ws.append(["Parâmetro", "Valor"])
            ws.append(["TMA (%)", self.tma])
            ws.append(["IR (%)", self.ir])
            ws.append(["CSLL (%)", self.csll])
            ws.append(["Taxa Efetiva (%)", self.get_effective_tax_rate()])
            ws.append(["Última Modificação", self.modified_at.strftime("%d/%m/%Y %H:%M:%S")])

            return True, "Configurações exportadas com sucesso"
            
        except Exception as e:
            return False, f"Erro ao exportar configurações: {str(e)}"

    def load_from_excel(self, workbook, sheet_name="Configuração"):
        """
        Load configuration from an Excel worksheet.
        
        Args:
            workbook: openpyxl workbook object
            sheet_name (str): Name of the worksheet
            
        Returns:
            tuple: (bool, str) - (success, message)
        """
        try:
            if sheet_name not in workbook.sheetnames:
                return False, "Planilha de configuração não encontrada"

            ws = workbook[sheet_name]
            data = {}

            # Read values (assuming specific row positions)
            for row in ws.iter_rows(min_row=2, max_row=4, values_only=True):
                if row[0] == "TMA (%)":
                    data["tma"] = float(row[1] or 0)
                elif row[0] == "IR (%)":
                    data["ir"] = float(row[1] or 0)
                elif row[0] == "CSLL (%)":
                    data["csll"] = float(row[1] or 0)

            return self.from_dict(data)
            
        except Exception as e:
            return False, f"Erro ao importar configurações: {str(e)}"

    def __str__(self):
        """String representation of the configuration."""
        return (f"Configuração:\n"
                f"  TMA: {self.tma:.2f}%\n"
                f"  IR: {self.ir:.2f}%\n"
                f"  CSLL: {self.csll:.2f}%\n"
                f"  Taxa Efetiva: {self.get_effective_tax_rate():.2f}%")
