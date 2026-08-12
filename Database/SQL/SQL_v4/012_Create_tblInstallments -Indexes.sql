CREATE INDEX IX_tblInstallments_PolicyID
ON tblInstallments (PolicyID);
CREATE INDEX IX_tblInstallments_DueDate
ON tblInstallments (DueDate);
CREATE INDEX IX_tblInstallments_Status
ON tblInstallments (Status);
CREATE INDEX IX_tblInstallments_PaymentDate
ON tblInstallments (PaymentDate);
CREATE UNIQUE INDEX IX_tblInstallments_PolicyNo
ON tblInstallments (PolicyID, InstallmentNo);
