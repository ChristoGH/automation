select 
policy.umaid,
claim.claimnumber as 'Claim Number',
claim.createddate as 'Created Date',
claim.createduser as 'Created User',
claim.incidentdate as 'Date of Loss',
claim.notificationdate as 'Notification Date',
policy.inceptiondate as 'Policy Inception Date',
client.clientname as 'Principal Member',
productsetup.productsetupname,
policy.concatdistinctactiveoptionname,
claim.manualstatus as 'ManualStatus',
claim.decisionstatus as 'DecisionProgress',
claim.statusalias as 'Claim Status',
SUM(claiminsureditem.reserve) 'Original Reserve',
SUM(claiminsureditem.calculatednetreserve - claiminsureditem.previouslypaid) as 'Current Reserve',
SUM(claiminsureditem.previouslypaid) as 'Total Payments',
policygroup.policygroupname as PolicyGroupName,
if(policy.policygroupid = '' or policygroup.policygroupname like '%IND%','PRIVATE','CORP') as 'PolicyGroup',
/*policygroup.legaltype as 'PolicyGroupType',*/
Year(claim.notificationdate) as 'Notification Year',
Month(claim.notificationdate) as 'Notification Month',
claim.decisionrejectionreason, 
claim.decisionrejectionnotes

from claim,policy,policygroup,client,claiminsureditem,productsetup

where 1=1 
and claim.notificationdate > '2020/09/30'
and policy.umaid = 'CEN'
group by claim.claimnumber
order by claim.claimnumber asc


select 
claim.claimnumber as 'Claim Number',
claim.createddate as 'Created Date',
claim.createduser as 'Created User',
claim.incidentdate as 'Date of Loss',
claim.notificationdate as 'Notification Date',
policy.inceptiondate as 'Policy Inception Date',
client.clientname as 'Principal Member',
productsetup.productsetupname,
policy.concatdistinctactiveoptionname,
claim.manualstatus as 'Status',
claim.decisionstatus as 'Progress',
SUM(claiminsureditem.reserve) 'Original Reserve',
SUM(claiminsureditem.calculatednetreserve - claiminsureditem.previouslypaid) as 'Current Reserve',
SUM(claiminsureditem.previouslypaid) as 'Total Payments',
policygroup.policygroupname as 'PolicyGroup'

from claim,policy,policygroup,client,claiminsureditem,productsetup

where 1=1
group by claim.claimnumber
order by claim.claimnumber asc


