CREATE TABLE [dbo].[Storage.Assettable] (
    [UserFirst]    TEXT       NOT NULL,
    [UserLast]     TEXT       NOT NULL,
    [Type]         TEXT       NOT NULL,
    [Asset]        INT        NOT NULL,
    [Description]  TEXT       NOT NULL,
    [PO]           INT        NOT NULL,
    [Department]   INT        NULL,
    [Serial]       NCHAR (20) NOT NULL,
    [TransferFrom] NCHAR (10) NULL,
    [TransferTo]   NCHAR (10) NULL,
    [Date]         NCHAR (10) NULL,
    PRIMARY KEY CLUSTERED ([Asset] ASC)
);

