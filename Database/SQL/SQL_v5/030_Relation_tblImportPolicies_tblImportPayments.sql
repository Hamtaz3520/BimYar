ALTER TABLE tblImportPayments
ADD CONSTRAINT FK_tblImportPayments_tblImportPolicies
FOREIGN KEY (ImportPolicyID)
REFERENCES tblImportPolicies (ImportPolicyID);
