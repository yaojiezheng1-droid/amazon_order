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
        self.root.geometry("1000x800")
        
        # Load product data
        self.products = self._load_products()
        
        # Initialize product pool/cart
        self.product_pool = {}  # {sku: {'product': product_dict, 'quantity': int}}
        
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
        
        # Quantity input and add button
        quantity_frame = ttk.Frame(product_frame)
        quantity_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Label(quantity_frame, text="Quantity:").pack(side=tk.LEFT, padx=(0, 10))
        
        self.quantity_var = tk.StringVar(value="1")
        quantity_spinbox = ttk.Spinbox(quantity_frame, from_=1, to=100000, 
                                     textvariable=self.quantity_var, width=10)
        quantity_spinbox.pack(side=tk.LEFT, padx=(0, 20))
        
        # Add to pool button
        ttk.Button(quantity_frame, text="Add to Pool", 
                  command=self._add_to_pool).pack(side=tk.LEFT, padx=(0, 10))
        
        # Update quantity button (for existing items)
        ttk.Button(quantity_frame, text="Update Quantity", 
                  command=self._update_quantity).pack(side=tk.LEFT)
        
        # Product Pool section
        pool_frame = ttk.LabelFrame(main_frame, text="Product Pool", padding="10")
        pool_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        pool_frame.columnconfigure(0, weight=1)
        pool_frame.rowconfigure(0, weight=1)
        
        # Create Treeview for product pool
        columns = ('SKU', 'Product Name', 'Quantity', 'Unit Price', 'Total Price')
        self.pool_tree = ttk.Treeview(pool_frame, columns=columns, show='headings', height=8)
        
        # Define headings
        for col in columns:
            self.pool_tree.heading(col, text=col)
            
        # Configure column widths
        self.pool_tree.column('SKU', width=120)
        self.pool_tree.column('Product Name', width=300)
        self.pool_tree.column('Quantity', width=80)
        self.pool_tree.column('Unit Price', width=100)
        self.pool_tree.column('Total Price', width=100)
        
        # Add scrollbar to treeview
        pool_scrollbar = ttk.Scrollbar(pool_frame, orient=tk.VERTICAL, command=self.pool_tree.yview)
        self.pool_tree.configure(yscrollcommand=pool_scrollbar.set)
        
        self.pool_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        pool_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Pool control buttons
        pool_button_frame = ttk.Frame(pool_frame)
        pool_button_frame.grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        
        ttk.Button(pool_button_frame, text="Remove Selected", 
                  command=self._remove_from_pool).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(pool_button_frame, text="Clear All", 
                  command=self._clear_pool).pack(side=tk.LEFT, padx=(0, 10))
        
        # Order name input
        ttk.Label(pool_button_frame, text="Order Name:").pack(side=tk.LEFT, padx=(20, 5))
        self.order_name_var = tk.StringVar(value="factory")
        order_name_entry = ttk.Entry(pool_button_frame, textvariable=self.order_name_var, width=15)
        order_name_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(pool_button_frame, text="Generate Command", 
                  command=self._generate_command).pack(side=tk.LEFT, padx=(10, 0))
        
        # Command output section
        command_frame = ttk.LabelFrame(main_frame, text="Generated Command", padding="10")
        command_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
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
        status_bar.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
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
        
        # If product is already in pool, show its current quantity
        if product["sku"] in self.product_pool:
            current_qty = self.product_pool[product["sku"]]["quantity"]
            self.quantity_var.set(str(current_qty))
            self.status_var.set(f"Selected: {product['sku']} - {product['name']} (Currently in pool: {current_qty})")
        else:
            self.quantity_var.set("1")
            self.status_var.set(f"Selected: {product['sku']} - {product['name']}")
    
    def _add_to_pool(self):
        """Add selected product to the pool"""
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
        
        # Add or update product in pool
        self.product_pool[sku] = {
            "product": self.selected_product,
            "quantity": quantity
        }
        
        self._refresh_pool_display()
        self.status_var.set(f"Added {quantity} × {sku} to pool")
    
    def _update_quantity(self):
        """Update quantity for existing product in pool"""
        if not hasattr(self, 'selected_product'):
            messagebox.showwarning("Warning", "Please select a product first")
            return
            
        sku = self.selected_product["sku"]
        if sku not in self.product_pool:
            messagebox.showwarning("Warning", f"Product {sku} is not in the pool. Use 'Add to Pool' instead.")
            return
            
        try:
            quantity = int(self.quantity_var.get())
            if quantity <= 0:
                raise ValueError("Quantity must be positive")
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid quantity: {e}")
            return
        
        self.product_pool[sku]["quantity"] = quantity
        self._refresh_pool_display()
        self.status_var.set(f"Updated {sku} quantity to {quantity}")
    
    def _refresh_pool_display(self):
        """Refresh the product pool display"""
        # Clear existing items
        for item in self.pool_tree.get_children():
            self.pool_tree.delete(item)
        
        # Add current pool items
        total_value = 0
        for sku, item in self.product_pool.items():
            product = item["product"]
            quantity = item["quantity"]
            unit_price = product["price"]
            total_price = unit_price * quantity
            total_value += total_price
            
            self.pool_tree.insert('', 'end', values=(
                sku,
                product["name"][:50] + ("..." if len(product["name"]) > 50 else ""),
                quantity,
                f"¥{unit_price}",
                f"¥{total_price:,.2f}"
            ))
        
        # Update frame title with count and total
        pool_count = len(self.product_pool)
        pool_frame = self.pool_tree.master
        pool_frame.configure(text=f"Product Pool ({pool_count} items, Total: ¥{total_value:,.2f})")
    
    def _remove_from_pool(self):
        """Remove selected item from pool"""
        selection = self.pool_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an item to remove")
            return
            
        # Get SKU from selected item
        item = self.pool_tree.item(selection[0])
        sku = item['values'][0]
        
        # Remove from pool
        if sku in self.product_pool:
            del self.product_pool[sku]
            self._refresh_pool_display()
            self.status_var.set(f"Removed {sku} from pool")
    
    def _clear_pool(self):
        """Clear all items from pool"""
        if not self.product_pool:
            messagebox.showinfo("Info", "Pool is already empty")
            return
            
        result = messagebox.askyesno("Confirm", "Are you sure you want to clear all items from the pool?")
        if result:
            self.product_pool.clear()
            self._refresh_pool_display()
            self.status_var.set("Cleared all items from pool")
    
    def _generate_command(self):
        """Generate the direct_sku_to_json.py command for all products in pool"""
        if not self.product_pool:
            messagebox.showwarning("Warning", "Please add products to the pool first")
            return
        
        # Get order name
        order_name = self.order_name_var.get().strip()
        if not order_name:
            order_name = "factory"
        
        # Build command with all SKU-quantity pairs
        command_parts = ["python", "direct_sku_to_json.py", "--name", order_name]
        
        total_value = 0
        product_details = []
        
        for sku, item in self.product_pool.items():
            product = item["product"]
            quantity = item["quantity"]
            unit_price = product["price"]
            total_price = unit_price * quantity
            total_value += total_price
            
            command_parts.extend([sku, str(quantity)])
            product_details.append(f"- {sku}: {quantity} × ¥{unit_price} = ¥{total_price:,.2f} ({product['name']})")
        
        command = " ".join(command_parts)
        
        # Display command with details
        details = f"""Command to generate order for {len(self.product_pool)} products:

{command}

Product Details:
{chr(10).join(product_details)}

Order Summary:
- Order Name: {order_name}
- Total Products: {len(self.product_pool)}
- Total Items: {sum(item['quantity'] for item in self.product_pool.values())}
- Total Value: ¥{total_value:,.2f}

This command will:
1. Generate {order_name}-1.json, {order_name}-2.json, etc. (grouped by supplier/factory)
2. Automatically convert to Excel format ({order_name}-1.xlsx, {order_name}-2.xlsx, etc.)
3. Apply current date ({self._get_current_date()}) and delivery date calculations
4. Include all required accessories based on accessory mapping
"""
        
        self.command_text.delete(1.0, tk.END)
        self.command_text.insert(1.0, details)
        
        # Store command for copying/execution
        self.current_command = command
        
        total_items = sum(item['quantity'] for item in self.product_pool.values())
        self.status_var.set(f"Generated command for {len(self.product_pool)} products ({total_items} total items)")
    
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
