#!/usr/bin/env python3
"""
Product Search GUI for Amazon Order Generation

This script provides a graphical interface to search for products by 产品名称 (Product Name) 
or 产品编号 (Product Code/SKU) and generate commands for direct_sku_to_json.py.

Features:
- Search products by name or SKU with dropdown suggestions
- Input quantity for selected products
- Generate and display the command to run direct_sku_to_json.py
- Copy command to clipboard for easy execution
"""

import json
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import subprocess
import sys
from pathlib import Path
from typing import Dict, List, Tuple
import pyperclip  # For clipboard functionality

class ProductSearchGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Amazon Order Product Search")
        self.root.geometry("800x600")
        
        # Load product data
        self.products = self._load_products()
        
        # Create GUI elements
        self._create_widgets()
        
    def _load_products(self) -> List[Dict]:
        """Load all products from JSON templates"""
        template_dir = Path(__file__).resolve().parent / "json_template"
        products = []
        
        try:
            for json_file in template_dir.glob("*.json"):
                try:
                    with open(json_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    
                    for product in data.get("products", []):
                        product_info = {
                            "sku": product.get("产品编号", ""),
                            "name": product.get("产品名称", ""),
                            "description": product.get("描述", ""),
                            "price": product.get("单价", 0),
                            "file": json_file.stem
                        }
                        if product_info["sku"]:  # Only add if SKU exists
                            products.append(product_info)
                            
                except Exception as e:
                    print(f"Error reading {json_file}: {e}")
                    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load products: {e}")
            
        return products
    
    def _create_widgets(self):
        """Create all GUI widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Search section
        search_frame = ttk.LabelFrame(main_frame, text="Product Search", padding="10")
        search_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        search_frame.columnconfigure(1, weight=1)
        
        # Search type selection
        ttk.Label(search_frame, text="Search by:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.search_type = tk.StringVar(value="name")
        search_type_frame = ttk.Frame(search_frame)
        search_type_frame.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        ttk.Radiobutton(search_type_frame, text="Product Name (产品名称)", 
                       variable=self.search_type, value="name",
                       command=self._on_search_type_change).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Radiobutton(search_type_frame, text="Product Code (产品编号)", 
                       variable=self.search_type, value="sku",
                       command=self._on_search_type_change).pack(side=tk.LEFT)
        
        # Search entry with autocomplete
        ttk.Label(search_frame, text="Search:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self._on_search_change)
        
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=50)
        self.search_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Dropdown for suggestions
        self.suggestion_combo = ttk.Combobox(search_frame, width=70, state="readonly")
        self.suggestion_combo.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=(5, 0))
        self.suggestion_combo.bind('<<ComboboxSelected>>', self._on_suggestion_select)
        
        # Selected product section
        product_frame = ttk.LabelFrame(main_frame, text="Selected Product", padding="10")
        product_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        product_frame.columnconfigure(1, weight=1)
        
        # Product details
        ttk.Label(product_frame, text="SKU:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.sku_label = ttk.Label(product_frame, text="", foreground="blue")
        self.sku_label.grid(row=0, column=1, sticky=tk.W)
        
        ttk.Label(product_frame, text="Name:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
        self.name_label = ttk.Label(product_frame, text="", wraplength=600)
        self.name_label.grid(row=1, column=1, sticky=tk.W)
        
        ttk.Label(product_frame, text="Price:").grid(row=2, column=0, sticky=tk.W, padx=(0, 10))
        self.price_label = ttk.Label(product_frame, text="")
        self.price_label.grid(row=2, column=1, sticky=tk.W)
        
        # Quantity input
        quantity_frame = ttk.Frame(product_frame)
        quantity_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Label(quantity_frame, text="Quantity:").pack(side=tk.LEFT, padx=(0, 10))
        
        self.quantity_var = tk.StringVar(value="1")
        quantity_spinbox = ttk.Spinbox(quantity_frame, from_=1, to=100000, 
                                     textvariable=self.quantity_var, width=10)
        quantity_spinbox.pack(side=tk.LEFT, padx=(0, 20))
        
        # Generate command button
        ttk.Button(quantity_frame, text="Generate Command", 
                  command=self._generate_command).pack(side=tk.LEFT, padx=(20, 0))
        
        # Command output section
        command_frame = ttk.LabelFrame(main_frame, text="Generated Command", padding="10")
        command_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        command_frame.columnconfigure(0, weight=1)
        command_frame.rowconfigure(1, weight=1)
        
        # Command display
        self.command_text = scrolledtext.ScrolledText(command_frame, height=6, wrap=tk.WORD)
        self.command_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Buttons
        button_frame = ttk.Frame(command_frame)
        button_frame.grid(row=1, column=0, sticky=tk.W)
        
        ttk.Button(button_frame, text="Copy to Clipboard", 
                  command=self._copy_command).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Execute Command", 
                  command=self._execute_command).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Clear", 
                  command=self._clear_command).pack(side=tk.LEFT)
        
        # Status bar
        self.status_var = tk.StringVar(value=f"Loaded {len(self.products)} products")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Initialize suggestions
        self._update_suggestions()
    
    def _on_search_type_change(self):
        """Handle search type change"""
        self.search_var.set("")
        self._update_suggestions()
        
    def _on_search_change(self, *args):
        """Handle search text change"""
        self._update_suggestions()
        
    def _update_suggestions(self):
        """Update the suggestion dropdown based on search text and type"""
        search_text = self.search_var.get().lower()
        search_type = self.search_type.get()
        
        if search_type == "name":
            filtered = [p for p in self.products 
                       if search_text in p["name"].lower()]
            suggestions = [f"{p['name']} ({p['sku']})" for p in filtered[:20]]
        else:  # sku
            filtered = [p for p in self.products 
                       if search_text in p["sku"].lower()]
            suggestions = [f"{p['sku']} - {p['name']}" for p in filtered[:20]]
        
        self.suggestion_combo['values'] = suggestions
        if suggestions and not search_text:
            self.suggestion_combo.set('')
    
    def _on_suggestion_select(self, event):
        """Handle suggestion selection"""
        selection = self.suggestion_combo.get()
        if not selection:
            return
            
        # Extract SKU from selection
        if self.search_type.get() == "name":
            # Format: "Product Name (SKU)"
            sku = selection.split('(')[-1].rstrip(')')
        else:
            # Format: "SKU - Product Name"
            sku = selection.split(' - ')[0]
        
        # Find the product
        product = next((p for p in self.products if p["sku"] == sku), None)
        if product:
            self._display_product(product)
    
    def _display_product(self, product):
        """Display selected product details"""
        self.selected_product = product
        self.sku_label.config(text=product["sku"])
        self.name_label.config(text=product["name"])
        self.price_label.config(text=f"¥{product['price']}")
        
        self.status_var.set(f"Selected: {product['sku']} - {product['name']}")
    
    def _generate_command(self):
        """Generate the direct_sku_to_json.py command"""
        if not hasattr(self, 'selected_product'):
            messagebox.showwarning("Warning", "Please select a product first")
            return
            
        try:
            quantity = int(self.quantity_var.get())
            if quantity <= 0:
                raise ValueError("Quantity must be positive")
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid quantity: {e}")
            return
        
        sku = self.selected_product["sku"]
        command = f"python direct_sku_to_json.py {sku} {quantity}"
        
        # Display command with details
        details = f"""Command to generate order for {quantity} pieces of {sku}:

{command}

Product Details:
- SKU: {sku}
- Name: {self.selected_product['name']}
- Price: ¥{self.selected_product['price']}
- Total Value: ¥{self.selected_product['price'] * quantity:,.2f}

This command will:
1. Generate factory JSON files with accessories
2. Automatically convert to Excel format
3. Apply current date ({self._get_current_date()}) and delivery date calculations
"""
        
        self.command_text.delete(1.0, tk.END)
        self.command_text.insert(1.0, details)
        
        # Store command for copying/execution
        self.current_command = command
        
        self.status_var.set(f"Generated command for {quantity} × {sku}")
    
    def _get_current_date(self):
        """Get current date in Chinese format"""
        from datetime import datetime
        return datetime.now().strftime('%Y年%m月%d日')
    
    def _copy_command(self):
        """Copy command to clipboard"""
        if hasattr(self, 'current_command'):
            try:
                pyperclip.copy(self.current_command)
                self.status_var.set("Command copied to clipboard")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to copy to clipboard: {e}")
        else:
            messagebox.showwarning("Warning", "No command to copy")
    
    def _execute_command(self):
        """Execute the generated command"""
        if not hasattr(self, 'current_command'):
            messagebox.showwarning("Warning", "No command to execute")
            return
            
        try:
            # Change to the script directory
            script_dir = Path(__file__).resolve().parent
            
            # Execute command
            result = subprocess.run(
                self.current_command.split(),
                cwd=script_dir,
                capture_output=True,
                text=True,
                check=True
            )
            
            # Show success message with output
            success_msg = f"Command executed successfully!\n\nOutput:\n{result.stdout}"
            if result.stderr:
                success_msg += f"\n\nWarnings:\n{result.stderr}"
                
            messagebox.showinfo("Success", success_msg)
            self.status_var.set("Command executed successfully")
            
        except subprocess.CalledProcessError as e:
            error_msg = f"Command failed with exit code {e.returncode}\n\nError:\n{e.stderr}\n\nOutput:\n{e.stdout}"
            messagebox.showerror("Execution Error", error_msg)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to execute command: {e}")
    
    def _clear_command(self):
        """Clear the command display"""
        self.command_text.delete(1.0, tk.END)
        if hasattr(self, 'current_command'):
            delattr(self, 'current_command')
        self.status_var.set("Command cleared")


def main():
    """Main function to run the GUI"""
    try:
        # Check if pyperclip is available
        import pyperclip
    except ImportError:
        print("Warning: pyperclip not installed. Clipboard functionality will not work.")
        print("Install it with: pip install pyperclip")
    
    root = tk.Tk()
    app = ProductSearchGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
