DECLARE @MyProj VARCHAR(MAX)
SET @MyProj=inputbox('Project Name:')
SELECT CLASSFCN SecurityMarking,
    CASE
        WHEN CONT_STATEMENT IS NULL THEN ''
        ELSE CONT_STATEMENT
    END DistributionStatement,
    CONVERT(varchar(10),STATUSDATE,126) ReportingPeriodEnDate,
    CONT_NAME ContractorName,
    CONT_IDTYPE ContractorIDCodeTypeID,
    CONT_IDCODE ContractorIDCode,
    ADDRESS ContractorAddress_Street,
    CITY ContractorAddress_City,
    STATE ContractorAddress_State,
    COUNTRY ContractorAddress_Country,
    ZIP ContractorAddress_ZipCode,
    CONT_REPN PointOfContactName,
    CONT_REPT PointOfContactTitle,
    CONT_REPPHONE PointOfContactTelephone,
    CONT_REPEMAIL PointOfContactEmail,
    CONTRACT ContractName,
    CONT_NO ContractNumber,
    CONT_TYPE ContractType,
    CONT_TASK ContractTaskOrEffortName,
    CONT_PROGRAM ProgramName,
    CONT_PHASE ProgramPhase,
    CASE 
        WHEN EVMS_ACC=1 THEN 'TRUE' 
        ELSE 'FALSE'
    END EVMSAccepted,
    CONVERT(varchar(10),EVMS_ADATE,126) EVMSAcceptanceDate
FROM PROGRAM
WHERE [PROGRAM]=@MyProj 