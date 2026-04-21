/**
 * ZATCA Engine - FATOORA Phase 2 Integration
 * Handles XML generation (UBL 2.1), Hashing, and Signing logic for Saudi Arabia.
 */

const ZatcaEngine = (() => {
  
  // ZATCA Cryptographic Constants
  const HASH_ALGO = 'SHA-256';

  /**
   * Generates a KSA-compliant UBL 2.1 XML for an invoice.
   * Following ZATCA specifications for Simplified Tax Invoice (B2C).
   */
  function generateUBLXML(invoice) {
    const today = new Date().toISOString().split('T')[0];
    const time = new Date().toISOString().split('T')[1].split('.')[0];
    const uuid = self.crypto.randomUUID ? self.crypto.randomUUID() : 'N/A';
    
    // Financial calculations
    const totalAmount = parseFloat(invoice.total || 0);
    const taxAmount = (totalAmount * 15 / 115).toFixed(2); // Assuming 15% VAT is inclusive
    const netAmount = (totalAmount - taxAmount).toFixed(2);

    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<Invoice xmlns="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2" 
         xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2" 
         xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2">
    <cbc:UBLVersionID>2.1</cbc:UBLVersionID>
    <cbc:CustomizationID>TRX-1.0</cbc:CustomizationID>
    <cbc:ProfileID>reporting:1.0</cbc:ProfileID>
    <cbc:ID>${invoice.id || 'INV-' + Date.now()}</cbc:ID>
    <cbc:UUID>${uuid}</cbc:UUID>
    <cbc:IssueDate>${invoice.date || today}</cbc:IssueDate>
    <cbc:IssueTime>${time}</cbc:IssueTime>
    <cbc:InvoiceTypeCode name="0211010">388</cbc:InvoiceTypeCode>
    <cbc:DocumentCurrencyCode>SAR</cbc:DocumentCurrencyCode>
    <cbc:TaxCurrencyCode>SAR</cbc:TaxCurrencyCode>
    
    <cac:AdditionalDocumentReference>
        <cbc:ID>ICV</cbc:ID>
        <cbc:UUID>1</cbc:UUID>
    </cac:AdditionalDocumentReference>
    
    <cac:AccountingSupplierParty>
        <cac:Party>
            <cac:PartyIdentification>
                <cbc:ID schemeID="CRN">1010101010</cbc:ID>
            </cac:PartyIdentification>
            <cac:PostalAddress>
                <cbc:StreetName>شارع الرياض الرئيس</cbc:StreetName>
                <cbc:CityName>الرياض</cbc:CityName>
                <cbc:PostalZone>12345</cbc:PostalZone>
                <cac:Country>
                    <cbc:IdentificationCode>SA</cbc:IdentificationCode>
                </cac:Country>
            </cac:PostalAddress>
            <cac:PartyTaxScheme>
                <cbc:CompanyID>300000000000003</cbc:CompanyID>
                <cac:TaxScheme>
                    <cbc:ID>VAT</cbc:ID>
                </cac:TaxScheme>
            </cac:PartyTaxScheme>
            <cac:PartyName>
                <cbc:Name>شركة مشروع الفروج الوطني</cbc:Name>
            </cac:PartyName>
        </cac:Party>
    </cac:AccountingSupplierParty>
    
    <cac:AccountingCustomerParty>
        <cac:Party>
            <cac:PartyName>
                <cbc:Name>${invoice.customer || 'عميل نقدي'}</cbc:Name>
            </cac:PartyName>
        </cac:Party>
    </cac:AccountingCustomerParty>
    
    <cac:TaxTotal>
        <cbc:TaxAmount currencyID="SAR">${taxAmount}</cbc:TaxAmount>
        <cac:TaxSubtotal>
            <cbc:TaxableAmount currencyID="SAR">${netAmount}</cbc:TaxableAmount>
            <cbc:TaxAmount currencyID="SAR">${taxAmount}</cbc:TaxAmount>
            <cac:TaxCategory>
                <cbc:ID>S</cbc:ID>
                <cbc:Percent>15.00</cbc:Percent>
                <cac:TaxScheme>
                    <cbc:ID>VAT</cbc:ID>
                </cac:TaxScheme>
            </cac:TaxCategory>
        </cac:TaxSubtotal>
    </cac:TaxTotal>
    
    <cac:LegalMonetaryTotal>
        <cbc:LineExtensionAmount currencyID="SAR">${netAmount}</cbc:LineExtensionAmount>
        <cbc:TaxExclusiveAmount currencyID="SAR">${netAmount}</cbc:TaxExclusiveAmount>
        <cbc:TaxInclusiveAmount currencyID="SAR">${totalAmount}</cbc:TaxInclusiveAmount>
        <cbc:PayableAmount currencyID="SAR">${totalAmount}</cbc:PayableAmount>
    </cac:LegalMonetaryTotal>
    
    <cac:InvoiceLine>
        <cbc:ID>1</cbc:ID>
        <cbc:InvoicedQuantity unitCode="PCE">${invoice.quantity || 1}</cbc:InvoicedQuantity>
        <cbc:LineExtensionAmount currencyID="SAR">${netAmount}</cbc:LineExtensionAmount>
        <cac:TaxTotal>
            <cbc:TaxAmount currencyID="SAR">${taxAmount}</cbc:TaxAmount>
        </cac:TaxTotal>
        <cac:Item>
            <cbc:Name>${invoice.product}</cbc:Name>
            <cac:ClassifiedTaxCategory>
                <cbc:ID>S</cbc:ID>
                <cbc:Percent>15.00</cbc:Percent>
                <cac:TaxScheme>
                    <cbc:ID>VAT</cbc:ID>
                </cac:TaxScheme>
            </cac:ClassifiedTaxCategory>
        </cac:Item>
        <cac:Price>
            <cbc:PriceAmount currencyID="SAR">${invoice.price}</cbc:PriceAmount>
        </cac:Price>
    </cac:InvoiceLine>
</Invoice>`;

    return xml;
  }

  /**
   * Generates a SHA-256 Hash of the XML content for Phase 2 integrity.
   */
  async function generateInvoiceHash(xmlString) {
    const encoder = new TextEncoder();
    const data = encoder.encode(xmlString);
    const hashBuffer = await crypto.subtle.digest(HASH_ALGO, data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    // ZATCA expects the hash in Base64
    const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
    // For manual viewing, we'll return both
    return {
      hex: hashHex,
      base64: btoa(String.fromCharCode(...new Uint8Array(hashBuffer)))
    };
  }

  /**
   * Main entry point to make an invoice ZATCA-Ready
   */
  async function prepareZatcaCompliantInvoice(invoice) {
    const xml = generateUBLXML(invoice);
    const hashResult = await generateInvoiceHash(xml);
    
    return {
      xml: xml,
      hash: hashResult.base64,
      hashHex: hashResult.hex,
      zatcaStatus: 'PREPARED', // Phase: Prepared -> Signed -> Reported
      version: 'UBL-2.1-KSA-1.0'
    };
  }

  return {
    prepareZatcaCompliantInvoice
  };

})();
