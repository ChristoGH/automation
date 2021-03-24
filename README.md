# automation
automating data workflows.
conda create --name automation python=3.7
conda install pandas

SELECT companyname,companynumber, policygroupname,policygroupid FROM policygroup
ORDER BY companyname DESC

SELECT DISTINCT 
policy.policynumber AS policy_policynumber,
policy.alternativepolicynumber AS policy_alternativepolicynumber,
policy.umaid AS policy_umaid,
productsetup.productsetupname AS productsetup_productsetupname,
policy.concatdistinctactiveoptionname as optionname,
policy.inceptiondate AS policy_inceptiondate,
policy.cancellationdate As policy_cancellationdate,
policy.status AS policy_status,
policy.createduser AS policy_createduser,
policy.createddate AS policy_createddate,
client.firstname AS client_firstname,
client.surname AS client_surname,
client.idnumber AS client_idnumber,
client.age AS client_age,
client.residentialaddress AS client_residentialaddress,
client.residentialsuburb AS client_residentialsuburb,
client.residentialcode AS client_residentialcode,
client.postaladdress AS client_postaladdress,
client.postalsuburb AS client_postalsuburb,
client.postalcode AS client_postalcode,
client.emailaddress AS client_emailaddress,
client.cellnumber AS client_cellnumber,
intermediarygroup.intermediarygroupname AS intermediarygroup_intermediarygroupname,
salesperson.salespersonname AS salesperson_salespersonname,
policy.paymentmethod AS policy_paymentmethod,
ifnull(policygroup.policygroupname,'Individual') AS policygroup_policygroupname,
policy.concatdistinctactiveoptionname AS productoptionsetup_productoptionname,
policy.grosspremium AS policy_grosspremium,
policy.totalfees AS policy_totalfees,
policy.totalpremium AS policy_totalpremium,
internalmarketer.internalmarketername AS internalmarketer_internalmarketername,
DATE_FORMAT(policy.createddate,'%Y/%m/01') as createdmonth,
CASE
	WHEN policy.status IN ('Lapsed','Cancelled')
	THEN (SELECT p.updateduser
		  FROM policy p
		  WHERE p.policynumber=policy.policynumber
		  AND p.previousstatus<>p.status
		  AND p.status IN ('Lapsed','Cancelled')
		  ORDER BY p.updateddate DESC
		  LIMIT 1)
		  
	ELSE ''
END as policy_cancelleduser

 
FROM policy,client,intermediarygroup,salesperson,policygroup,productsetup,internalmarketer


-- PolicyGroupInformation
SELECT
	policygroup.internalaccountmanager,
	policy.policynumber,
	policy.inceptiondate,
	client.clientname,
	client.idnumber,
	policy.policygroupid,
	policygroup.policygroupname,
	policy.paymentmethod,
	policy.directdebitstatus,
	policy.grosspremium,
	policy.totalpremium,
	policy.bankaccountnumber,
	policy.bankaccountholder,
	productsetup.productsetupname,
	productoptionsetup.productoptionname
FROM policy,policygroup,client,productsetup,productoptionsetup
WHERE IFNULL(policy.policygroupid,'')<>''
and policy.umaid='AUH'
and policy.status='Live'

SELECT * FROM policy

SELECT
	policygroup.internalaccountmanager,
	policy.policynumber,
	policy.inceptiondate,
	client.clientname,
	client.idnumber,
	policy.policygroupid,
	policygroup.policygroupname,
	policy.paymentmethod,
	policy.directdebitstatus,
	policy.grosspremium,
	policy.totalpremium,
	policy.bankaccountnumber,
	policy.bankaccountholder,
	productsetup.productsetupname,
	productoptionsetup.productoptionname
FROM policy,policygroup,client,productsetup,productoptionsetup
WHERE IFNULL(policy.policygroupid,'')<>''
and policy.umaid='AUH'
and policy.status='Live'

-- 
SELECT policygroupid,policygroupname,companynumber,companyname FROM policygroup ORDER BY policygroupname LIMIT 1000 
SELECT * FROM policygroupholdings LIMIT 20
productoptionsetup

No results found using the following criteria:
Query Name: = 'DetailedActiveMembersByPolicyGroup'
Description: = 'DetailedActiveMembersByPolicyGroup'
SQL: = 'select
productsetup.productsetupname,
policy.policynumber,
policy.status,
policy.policygroupid,
policy.policygroupholdingsid,
policy.inceptiondate,
policy.paymentmethod As PolicyPaymentMethod,
policygroup.policygroupname,
policygroup.paymentmethod As PolicyGroupPaymentMethod,
policy.grosspremium,
policy.outstandingbalance

/*policy.policynumber,
policy.status,
policy.policygroupid,
policygroupholdingsname,
policygroupholdings.alternativesystemid,
policygroup.policygroupid,
policygroupname*/
/*productsetupname,
productoptionname*/
/*intermediarygroup.intermediarygroupname, */
from
policy,
policygroup, productsetup
/*policyinsureditem,
policygroupholdings,
policygroup,
productsetup,
productoptionsetup
intermediarygroup*/
,productsetupname,productoptionname*/'



SELECT policygroupholdingsid,policygroupholdingsname,companyname FROM policygroupholdings ORDER BY policygroupholdingsname ASC LIMIT 200