using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Translima_Application
{
    [XmlRoot(ElementName = "createDate")]
    public class CreateDate
    {
        [XmlElement(ElementName = "dateInTimezone")]
        public string DateInTimezone { get; set; }
        [XmlElement(ElementName = "timezone")]
        public string Timezone { get; set; }
    }

    [XmlRoot(ElementName = "customsOffices")]
    public class CustomsOffices
    {
        [XmlElement(ElementName = "identCcode")]
        public string IdentCcode { get; set; }
        [XmlElement(ElementName = "officeType")]
        public string OfficeType { get; set; }
    }

    [XmlRoot(ElementName = "extraFields")]
    public class ExtraFields
    {
        [XmlElement(ElementName = "key")]
        public string Key { get; set; }
        [XmlElement(ElementName = "value")]
        public string Value { get; set; }
    }
    

    [XmlRoot(ElementName = "invoiceDate")]
    public class InvoiceDate
    {
        [XmlElement(ElementName = "dateInTimezone")]
        public string DateInTimezone { get; set; }
        [XmlElement(ElementName = "timezone")]
        public string Timezone { get; set; }
    }

    [XmlRoot(ElementName = "invoicePrice")]
    public class InvoicePrice
    {
        [XmlElement(ElementName = "value")]
        public string Value { get; set; }
        [XmlElement(ElementName = "currencyIso")]
        public string CurrencyIso { get; set; }
    }

    [XmlRoot(ElementName = "invoices")]
    public class Invoices
    {
        [XmlElement(ElementName = "invoiceDate")]
        public InvoiceDate InvoiceDate { get; set; }
        [XmlElement(ElementName = "invoiceNumber")]
        public string InvoiceNumber { get; set; }
        [XmlElement(ElementName = "invoicePrice")]
        public InvoicePrice InvoicePrice { get; set; }
    }

    [XmlRoot(ElementName = "grossWeight")]
    public class GrossWeight
    {
        [XmlElement(ElementName = "value")]
        public string Value { get; set; }
        [XmlElement(ElementName = "unit")]
        public string Unit { get; set; }
    }

    [XmlRoot(ElementName = "netPrice")]
    public class NetPrice
    {
        [XmlElement(ElementName = "value")]
        public string Value { get; set; }
        [XmlElement(ElementName = "currencyIso")]
        public string CurrencyIso { get; set; }
    }

    [XmlRoot(ElementName = "netWeight")]
    public class NetWeight
    {
        [XmlElement(ElementName = "value")]
        public string Value { get; set; }
        [XmlElement(ElementName = "unit")]
        public string Unit { get; set; }
    }

    public class AdditionalReferences
    {
        [XmlElement(ElementName = "issueDate")]
        public IssueDate IssueDate { get; set; }

        [XmlElement(ElementName = "itemNumber")]
        public string itemNumber { get; set; }

        [XmlElement(ElementName = "reference")]
        public string reference { get; set; }

        [XmlElement(ElementName = "referenceType")]
        public string referenceType { get; set; }
    }

    public class IssueDate
    {
        [XmlElement(ElementName = "dateInTimezone")]
        public string dateInTimezone { get; set; }
        [XmlElement(ElementName = "timezone")]
        public string timezone { get; set; }
    }

    [XmlRoot(ElementName = "packageGrossMass")]
    public class PackageGrossMass
    {
        [XmlElement(ElementName = "value")]
        public string Value { get; set; }
        [XmlElement(ElementName = "unit")]
        public string Unit { get; set; }
    }

    [XmlRoot(ElementName = "packages")]
    public class Packages
    {
        [XmlElement(ElementName = "marks")]
        public string Marks { get; set; }
        [XmlElement(ElementName = "number")]
        public string Number { get; set; }
        [XmlElement(ElementName = "packageType")]
        public string PackageType { get; set; }

        [XmlElement(ElementName = "packedItemQuantity")]
        public string packedItemQuantity { get; set; }
        [XmlElement(ElementName = "packageIdClientSystem")]
        public string packageIdClientSystem { get; set; }
        [XmlElement(ElementName = "packageGrossMass")]
        public PackageGrossMass PackageGrossMass { get; set; }
    }


    [XmlRoot(ElementName = "items")]
    public class Items
    {
        
        [XmlElement(ElementName = "extraFields")]
        public List<ExtraFields> ExtraFields { get; set; }
        [XmlElement(ElementName = "goodsDescription")]
        public string GoodsDescription { get; set; }
        [XmlElement(ElementName = "grossWeight")]
        public GrossWeight GrossWeight { get; set; }
        [XmlElement(ElementName = "invoiceNumbers")]
        public string InvoiceNumbers { get; set; }
        [XmlElement(ElementName = "itemNumber")]
        public string ItemNumber { get; set; }
        [XmlElement(ElementName = "materialNumber")]
        public string MaterialNumber { get; set; }
        [XmlElement(ElementName = "netPrice")]
        public NetPrice NetPrice { get; set; }
        [XmlElement(ElementName = "netWeight")]
        public NetWeight NetWeight { get; set; }
        [XmlElement(ElementName = "originCountryCode")]
        public string OriginCountryCode { get; set; }
        [XmlElement(ElementName = "packages")]
        public List<Packages> Packages { get; set; }
        [XmlElement(ElementName = "sequenceNumber")]
        public string SequenceNumber { get; set; }
        [XmlElement(ElementName = "quantity")]
        public Quantity Quantity { get; set; }
        [XmlElement(ElementName = "tariffNumber")]
        public string TariffNumber { get; set; }
        [XmlElement(ElementName = "preferentialCountryCode")]
        public string PreferentialCountryCode { get; set; }
        [XmlElement(ElementName = "statisticalValue")]
        public StatisticalValue StatisticalValue { get; set; }

        
    }

    [XmlRoot(ElementName = "statisticalValue")]
    public class StatisticalValue
    {
        [XmlElement(ElementName = "value")]
        public string Value { get; set; }
        [XmlElement(ElementName = "currencyIso")]
        public string CurrencyIso { get; set; }
    }

    [XmlRoot(ElementName = "quantity")]
    public class Quantity
    {
        [XmlElement(ElementName = "value")]
        public string Value { get; set; }
        [XmlElement(ElementName = "unit")]
        public string Unit { get; set; }
    }

    [XmlRoot(ElementName = "parties")]
    public class Parties
    {
        [XmlElement(ElementName = "city")]
        public string City { get; set; }
        [XmlElement(ElementName = "countryCode")]
        public string CountryCode { get; set; }
        [XmlElement(ElementName = "name")]
        public string Name { get; set; }
        [XmlElement(ElementName = "partyType")]
        public string PartyType { get; set; }
        [XmlElement(ElementName = "postalCode")]
        public string PostalCode { get; set; }
        [XmlElement(ElementName = "street")]
        public string Street { get; set; }
        [XmlElement(ElementName = "traderId")]
        public string TraderId { get; set; }
        [XmlElement(ElementName = "vatId")]
        public string VatId { get; set; }
        [XmlElement(ElementName = "companyCode")]
        public string CompanyCode { get; set; }
    }

    [XmlRoot(ElementName = "personInCharge")]
    public class PersonInCharge
    {
        [XmlElement(ElementName = "name")]
        public string Name { get; set; }
        [XmlElement(ElementName = "phone")]
        public string Phone { get; set; }
        [XmlElement(ElementName = "email")]
        public string Email { get; set; }
    }

    [XmlRoot(ElementName = "totalGrossWeight")]
    public class TotalGrossWeight
    {
        [XmlElement(ElementName = "value")]
        public string Value { get; set; }
        [XmlElement(ElementName = "unit")]
        public string Unit { get; set; }
    }

    [XmlRoot(ElementName = "totalInvoicePrice")]
    public class TotalInvoicePrice
    {
        [XmlElement(ElementName = "value")]
        public string Value { get; set; }
        [XmlElement(ElementName = "currencyIso")]
        public string CurrencyIso { get; set; }
    }

    [XmlRoot(ElementName = "totalNetWeight")]
    public class TotalNetWeight
    {
        [XmlElement(ElementName = "value")]
        public string Value { get; set; }
        [XmlElement(ElementName = "unit")]
        public string Unit { get; set; }
    }

    [XmlRoot(ElementName = "transportMeans")]
    public class TransportMeans
    {
        [XmlElement(ElementName = "identity")]
        public string Identity { get; set; }
        [XmlElement(ElementName = "meansType")]
        public string MeansType { get; set; }
        [XmlElement(ElementName = "modeType")]
        public string ModeType { get; set; }
        [XmlElement(ElementName = "nationality")]
        public string Nationality { get; set; }
    }
    [XmlRoot(ElementName = "amount")]
    public class costamount
    {
        [XmlElement(ElementName = "currencyIso")]
        public string currencyIsoamount { get; set; }
        [XmlElement(ElementName = "value")]
        public string valueamount { get; set; }

    }
    public class amount
    {
        [XmlElement(ElementName = "amount")]
        public costamount costsamount { get; set; }

        [XmlElement(ElementName = "type")]
        public string type { get; set; }
    }

    [XmlRoot(ElementName = "declaration")]
    public class Declaration
    {
        [XmlElement(ElementName = "commercialReference")]
        public string CommercialReference { get; set; }
        [XmlElement(ElementName = "customsOffices")]
        public List<CustomsOffices> CustomsOffices { get; set; }
        [XmlElement(ElementName = "additionalReferences")]
        public List<AdditionalReferences> AdditionalReferences { get; set; }
        [XmlElement(ElementName = "destinationCountryCode")]
        public string DestinationCountryCode { get; set; }
        [XmlElement(ElementName = "dispatchCountryCode")]
        public string DispatchCountryCode { get; set; }

        [XmlElement(ElementName = "costs")]
        public List<amount> costs { get; set; }

        [XmlElement(ElementName = "extraFields")]
        public List<ExtraFields> ExtraFields { get; set; }
        [XmlElement(ElementName = "goodsDescription")]
        public string GoodsDescription { get; set; }
        [XmlElement(ElementName = "incotermCode")]
        public string IncotermCode { get; set; }
        [XmlElement(ElementName = "incotermPlace")]
        public string IncotermPlace { get; set; }
        [XmlElement(ElementName = "instructions")]
        public string Instructions { get; set; }
        [XmlElement(ElementName = "invoices")]
        public List<Invoices> Invoices { get; set; }
        [XmlElement(ElementName = "items")] 
        public List<Items> Items { get; set; }
        [XmlElement(ElementName = "localReference")]
        public string LocalReference { get; set; }
        [XmlElement(ElementName = "parties")]
        public List<Parties> Parties { get; set; }
        [XmlElement(ElementName = "personInCharge")]
        public PersonInCharge PersonInCharge { get; set; }
        [XmlElement(ElementName = "totalGrossWeight")]
        public TotalGrossWeight TotalGrossWeight { get; set; }
        [XmlElement(ElementName = "totalInvoicePrice")]
        public TotalInvoicePrice TotalInvoicePrice { get; set; }
        [XmlElement(ElementName = "totalNetWeight")]
        public TotalNetWeight TotalNetWeight { get; set; }
        [XmlElement(ElementName = "transportMeans")]
        public TransportMeans TransportMeans { get; set; }
    }

    [XmlRoot(ElementName = "brokerFileDTO")]
    public class BrokerFileDTO
    {
        [XmlElement(ElementName = "clientCode")]
        public string ClientCode { get; set; }
        [XmlElement(ElementName = "createDate")]
        public CreateDate CreateDate { get; set; }
        [XmlElement(ElementName = "declaration")]
        public Declaration Declaration { get; set; }
        [XmlElement(ElementName = "declaringCountryCode")]
        public string DeclaringCountryCode { get; set; }
        [XmlElement(ElementName = "process")]
        public string Process { get; set; }
        [XmlElement(ElementName = "version")]
        public string Version { get; set; }
        

    }

}
