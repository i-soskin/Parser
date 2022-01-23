CREATE TABLE [dbo].[Table]
(
	[Id] INT NOT NULL PRIMARY KEY IDENTITY ,
	[Name] varchar NOT NULL,
	[Description] varchar NOT NULL,
	[Source] varchar NOT NULL,
	[Object] varchar NOT NULL,
	[Сonfidential] INT NOT NULL DEFAULT 0,
	[Integrity] INT NOT NULL DEFAULT 0,
	[Avalibilty] INT NOT NULL DEFAULT 0,
	[DateAdd] DATE NOT NULL,
	[DataChange] DATE NOT NULL

)
