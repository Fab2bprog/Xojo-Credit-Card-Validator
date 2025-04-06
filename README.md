# Xojo-Credit-Card-Validator
This simple Xojo application allows users to validate various types of financial identifiers commonly used around the world. It is designed as a learning tool and utility app to demonstrate basic number validation logic using Xojo’s object-oriented programming capabilities.

## ✅ Supported Validations

- **Credit Card Numbers**  
  Supports major card providers such as:
  - Visa  
  - Mastercard  
  - American Express...  
  Validation is based on the Luhn algorithm.

- **IBAN (International Bank Account Number)**  
  Checks the structure and checksum according to international IBAN rules.

- **U.S. Bank Account Numbers**  
  Performs structural validation for U.S. routing and account numbers.

---

## 🧱 Project Structure

The project is composed of **three reusable classes**:

- `CreditCardValidator`  
  Detects the card type and verifies number validity.

- `IBANValidator`  
  Verifies IBAN structure and checksum.

- `USBankAccountValidator`  
  Validates U.S. account and routing numbers format.

