from datetime import datetime
import uuid

class BaseFinancialItem:
    def __init__(self, tag=None, description="", quantity=0, unit_price=0.0):
        """
        Initialize a base financial item.
        
        Args:
            tag (str, optional): Unique identifier for the item. If None, generates a new one.
            description (str): Description of the item
            quantity (float): Quantity of the item
            unit_price (float): Unit price of the item
        """
        self.tag = tag if tag else self._generate_tag()
        self.description = description
        self.quantity = float(quantity)
        self.unit_price = float(unit_price)
        self.created_at = datetime.now()
        self.modified_at = datetime.now()

    def _generate_tag(self):
        """Generate a unique TAG for the item."""
        return f"ITEM_{str(uuid.uuid4())[:8]}"

    @property
    def total_value(self):
        """Calculate the total value of the item."""
        return self.quantity * self.unit_price

    def update(self, description=None, quantity=None, unit_price=None):
        """
        Update the item's attributes.
        
        Args:
            description (str, optional): New description
            quantity (float, optional): New quantity
            unit_price (float, optional): New unit price
        """
        if description is not None:
            self.description = description
        if quantity is not None:
            self.quantity = float(quantity)
        if unit_price is not None:
            self.unit_price = float(unit_price)
        self.modified_at = datetime.now()

    def to_dict(self):
        """
        Convert the item to a dictionary.
        
        Returns:
            dict: Dictionary representation of the item
        """
        return {
            "tag": self.tag,
            "description": self.description,
            "quantity": self.quantity,
            "unit_price": self.unit_price,
            "total_value": self.total_value,
            "created_at": self.created_at.isoformat(),
            "modified_at": self.modified_at.isoformat()
        }

    def to_row(self):
        """
        Convert the item to a row format for Excel.
        
        Returns:
            list: List containing item data in order [TAG, Description, Quantity, Unit Price, Total Value]
        """
        return [
            self.tag,
            self.description,
            self.quantity,
            self.unit_price,
            self.total_value
        ]

    @classmethod
    def from_dict(cls, data):
        """
        Create an item from a dictionary.
        
        Args:
            data (dict): Dictionary containing item data
            
        Returns:
            BaseFinancialItem: New instance of the item
        """
        item = cls(
            tag=data.get("tag"),
            description=data.get("description", ""),
            quantity=data.get("quantity", 0),
            unit_price=data.get("unit_price", 0.0)
        )
        if "created_at" in data:
            item.created_at = datetime.fromisoformat(data["created_at"])
        if "modified_at" in data:
            item.modified_at = datetime.fromisoformat(data["modified_at"])
        return item

    def validate(self):
        """
        Validate the item's data.
        
        Returns:
            tuple: (bool, str) - (is_valid, error_message)
        """
        if not self.tag:
            return False, "TAG não pode estar vazio"
        if not self.description.strip():
            return False, "Descrição não pode estar vazia"
        if self.quantity < 0:
            return False, "Quantidade não pode ser negativa"
        if self.unit_price < 0:
            return False, "Valor unitário não pode ser negativo"
        return True, ""

    def __str__(self):
        """String representation of the item."""
        return f"{self.tag} - {self.description} ({self.quantity} x R${self.unit_price:.2f} = R${self.total_value:.2f})"

    def __repr__(self):
        """Detailed string representation of the item."""
        return (f"<{self.__class__.__name__}("
                f"tag='{self.tag}', "
                f"description='{self.description}', "
                f"quantity={self.quantity}, "
                f"unit_price={self.unit_price})>")
