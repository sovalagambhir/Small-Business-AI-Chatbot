# Small-Business-AI-Chatbot
# Adarsh Garments AI Assistant Desktop Management System

An efficient, secure, and localized desktop management application specifically designed for small retail storefront environments. This system provides automated customer support interactions, dynamic information retrieval, transactional order logging, and secure administrative alert channels without requiring expensive external cloud hosting or third-party token dependencies.

---

## 🛠️ Core Architecture & Core Technologies

The software is engineered entirely within a lightweight, local environment to ensure offline resilience and zero operating costs:
* **User Interface Layer:** Native Python `Tkinter` canvas implementing high-contrast visual blocks.
* **Conversational Logic:** Typo-tolerant string comparison engine powered by the `difflib.SequenceMatcher` algorithm.
* **Persistent Storage Hub:** Flat-file data architecture leveraging local JSON schemas and automated `OpenPyXL` spreadsheet ledgers.
* **Notification Tunnel:** Asynchronous multi-threaded background workers utilizing secure `smtplib` TLS relays.

---

## 📂 System File Architecture & Data Ecosystem

Upon initial execution, the application runs a built-in environment health validation routine. If missing, it will automatically heal and generate the following local database file matrices:

* `orders.xlsx`: The central database containing columns for Order ID, Customer Name, Contact Email, Product SKU, Shipping Address, Fulfillment Status, and Transactional System Notes.
* `intents.json`: Houses the conversational dialogue profiles, pattern matching matrix blocks, and variant response scripts.
* `discount.txt`: Stores active shop-floor voucher strings and immediate flat markdown promotions.
* `current_offers.txt`: Retains multi-item seasonal campaign details and in-store clearance rules.

---

## 🚀 Deployment & Local Installation Guide

### 1. Environmental Prerequisites
Ensure your terminal environment is running **Python 3.10** or a more recent stable release. 

### 2. Clone and Initialize the Working Directory
Open your terminal window and navigate to your production script folder:
```bash
cd path/to/your/project-folder
