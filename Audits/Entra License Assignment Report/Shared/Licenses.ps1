function Get-LicenseFriendlyName {
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]$Sku
    )


    Switch ($Sku) {
        "AAD_PREMIUM" {  "Azure Active Directory Premium" }
        "AAD_PREMIUM_P2" {  "Azure Active Directory Premium 2" }
        "ATP_ENTERPRISE" {  "Office 365 Advanced Threat Protection (Plan 1)" }
        "AX7_USER_TRIAL" {  "Microsoft Dynamics AX7 User Trial" }
        "CCIBOTS_PRIVPREV_VIRAL" {  "Power Virtual Agents Viral Trial" }
        "D365_SALES_PRO" {  "Dynamics 365 Sales Professional" }
        "D365_CUSTOMER_SERVICE_PRO_ATTACH" {  "Dynamics 365 Customer Service" }
        "DESKLESSPACK" {  "Office 365 F3" }
        "Dynamics_365_Customer_Service_Enterprise_viral_" {  "Dynamics_365_Customer_Service_Enterprise_viral_" }
        "DYN365_CUSTOMER_SERVICE_PRO" {  "Dynamics 365 Customer Service Professional" }
        "D365_CUSTOMER_SERVICE_PRO_ATTACH" {  "Dynamics 365 Customer Service" }
        "DYN365_ENTERPRISE_P1_IW" {  "Dynamics 365 P1 Trial for Information Workers" }
        "DYN365_ENTERPRISE_PLAN1" {  "Dynamics 365 Customer Engagement Plan 1" }
        "DYN365_ENTERPRISE_PLAN2" {  "Dynamics 365 Customer Engagement Plan 2" }
        "D365_CUSTOMER_SERVICE_ENT_ATTACH" { "Dynamics 365 Customer Service Enterprise Attach to Qualifying Dynamics 365 Base Offer"}
        "DYN365_CS_VOICE" { "Dynamics 365 Customer Service Voice Channel Add-in"}
        "DYN365_ENTERPRISE_SALES" { "Dynamics 365 Sales Enterprise Edition"}
        "D365_SALES_ENT_ATTACH" { "Dynamics 365 Sales Enterprise Attach to Qualifying Dynamics 365 Base Offer"}
        "DYN365_FINANCE" { "Dynamics 365 Finance"}
        "DYN365_ENTERPRISE_CUSTOMER_SERVICE" { "Dynamics 365 for Customer Service Enterprise Edition"}
        "DYN365_TEAM_MEMBERS" { "Dynamics 365 Team Members_wDynamicsRetail"}
        "DYN365_SCM" { "Dynamics 365 for Supply Chain Management"}
        "DYN365_PROJECT_OPERATIONS" { "Dynamics 365 Project Operations"}
        "DYN365_SCM_ATTACH" { "Dynamics 365 for Supply Chain Management Attach to Qualifying Dynamics 365 Base Offer"}
        "Dynamics_365_Sales_Premium_Viral_Trial" { "Dynamics 365 Sales Premium Viral Trial"}
        "Dynamics_365_Customer_Service_Enterprise_viral_trial" { "Dynamics 365 Customer Service Enterprise Viral Trial"}
        "EMS" {  "ENTERPRISE MOBILITY + SECURITY E3" }
        "ENTERPRISEPACK" {  "Office 365 E3" }
        "ENTERPRISEPREMIUM_NOPSTNCONF" {  "Office 365 E5 No PSTN" }
        "EXCHANGEARCHIVE" {  "EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER" }
        "EXCHANGEENTERPRISE" {  "EXCHANGE ONLINE (PLAN 2)" }
        "EXCHANGESTANDARD" {  "EXCHANGE ONLINE (PLAN 1)" }
        "FLOW_FREE" {  "Power Automate" }
        "FLOW_PER_USER" {  "Power Automate per user plan" }
        "FORMS_PRO" {  "Dynamics 365 Customer Voice Trial" }
        "IDENTITY_THREAT_PROTECTION" {  "Microsoft 365 E5 Security" }
        "M365_F1_COMM" {  "Microsoft 365 F1" }
        "Microsoft_365_Copilot" {"Copilot for Microsoft 365"}
        "M365EDU_A3_FACULTY" {"Microsoft 365 A3 for faculty"}
        "M365EDU_A3_STUUSEBNFT" {"Microsoft 365 A3 student use benefits"}
        "MCOMEETADV" {  "Skype for Business PSTN Conferencing" }
        "MCOEV" { "Microsoft Teams Phone Standard"}
        "MDATP_XPLAT" {  "Microsoft Defender for Endpoint" }
        "MEETING_ROOM" {  "Microsoft Teams Rooms Standard" }
        "Microsoft_Entra_ID_Governance_Step_Up_for_Microsoft_Entra_ID_P2" {"Microsoft Entra ID Governance Step-Up Entra ID 2"}
        "MICROSOFT_ENTRA_ID_GOVERNANCE_STEP_UP_FOR_MICROSOFT_ENTRA_ID_P2_FOR_FLW" {"Microsoft Entra ID Governance Step-Up Entra ID 2 Front Line Workers"}
        "Microsoft_Teams_Rooms_Pro" {  "Microsoft Teams Rooms Pro" }
        "Microsoft_Teams_Rooms_Basic" {"Microsoft Teams Rooms Basic"}
        "Microsoft_Teams_Premium" {"Microsoft Teams Premium Introductory Pricing"}
        "Microsoft_Teams_EEA_New" {"Microsoft Teams EEA"}
        "Microsoft_365_F1_EEA_(no_Teams)" {"Microsoft 365 F1 EEA (no Teams)"}
        "Microsoft_365_F3_EEA_(no_Teams)" {"Microsoft 365 F3 EEA (no Teams)"}
        "MTR_PREM" {  "Teams Rooms Premium" }
        "NONPROFIT_PORTAL" {  "NONPROFIT PORTAL" }
        "O365_BUSINESS_ESSENTIALS" {  "MICROSOFT 365 BUSINESS BASIC" }
        "O365_BUSINESS_PREMIUM" {  "MICROSOFT 365 BUSINESS STANDARD" }
        "O365_w/o Teams Bundle_M3" { "Microsoft 365 E3 EEA (no Teams)"}
        "PBI_PREMIUM_PER_USER_ADDON" {  "Power BI Premium Per User Add-On" }
        "POWER_BI_PRO" {  "Power BI Pro" }
        "POWER_BI_STANDARD" {  "Power BI" }
        "POWERAPPS_PER_USER" {  "Power Apps per User" }
        "POWERAPPS_VIRAL" {  "Microsoft Power Apps Plan 2 Trial" }
        "POWERAPPS_DEV" {"Microsoft PowerApps for Developer"}
        "POWERAUTOMATE_ATTENDED_RPA" {  "Power Automate per user with attended RPA plan" }
        "Power_Pages_vTrial_for_Makers" {  "Power Pages vTrial for Makers" }
        "PROJECT_P1" {"Project Plan 1"}
        "PROJECT_MADEIRA_PREVIEW_IW_SKU" {  "Dynamics 365 for Financials for IWs" }
        "PROJECTPROFESSIONAL" {  "PROJECT ONLINE PROFESSIONAL" }
        "PROJECT_PLAN3_DEPT" {  "Project Plan 3 (for Department)" }
        "RIGHTSMANAGEMENT_ADHOC" {  "Azure Rights Management Services Ad-hoc" }
        "SMB_APPS" {  "Business Apps" }
        "SPB" {  "MICROSOFT 365 BUSINESS PREMIUM" }
        "SPE_E3" {  "Microsoft 365 E3" }
        "SPE_E5" {  "Microsoft 365 E5" }
        "SPE_F3" {  "Microsoft 365 F3" }
        "SPE_F1" {  "Microsoft 365 F3" }
        "SPE_F5_SEC" {  "Microsoft 365 F5 Security Add-on" }
        "STANDARDPACK" {  "Office 365 E1" }
        "STANDARDWOFFPACK_STUDENT" {"Office 365 A1 for Students"}
        "STREAM" {  "Microsoft Stream" }
        "TEAMS_EXPLORATORY" {  "Microsoft Teams Exploratory" }
        "VISIOCLIENT" {  "VISIO Online Plan 2" }
        "VIRTUAL_AGENT_USL" {  "Microsoft Copilot Studio User License" }
        "WIN_DEF_ATP" {  "Windows Defender ATP" }
        "WINDOWS_STORE" {  "WINDOWS STORE FOR BUSINESS" }
        "E3_VDA_only" {  "Windows 10/11 Enterprise VDA" }
        "THREAT_INTELLIGENCE" {  "Microsoft Defender for Office 365 (Plan 2)" }
        "EMSPREMIUM" {  "Enterprise Mobility + Security E5" }
        "POWER_BI_PRO_DEPT" {  "Power BI Pro Dept" }
        "DYN365_BUSINESS_MARKETING" {  "Dynamics 365 for Marketing Business Edition" }


        "" {  "Unlicensed" }


        default {  "Unknown license" }
    } #End Switch

}