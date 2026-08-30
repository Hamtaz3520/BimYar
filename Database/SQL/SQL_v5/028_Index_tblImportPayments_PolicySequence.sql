CREATE UNIQUE INDEX UX_tblImportPayments_PolicySequence
ON tblImportPayments (ImportPolicyID, PaymentSequence);
