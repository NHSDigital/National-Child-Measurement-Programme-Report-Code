SELECT [GEOGRAPHY_CODE] as OrgCode
      ,[GEOGRAPHY_NAME] as OrgName
      ,[PARENT_GEOGRAPHY_CODE] as ParentCode
      ,[DATE_OF_OPERATION] as OpenDate
      ,*
  FROM [DATABASE].[SERVER].[TABLE]
  WHERE [ENTITY_CODE] in ('E06','E07','E08','E09','E10','E12','E92')
  AND ([DATE_OF_OPERATION] <= '<FYEnd>' AND ([DATE_OF_TERMINATION] IS NULL OR [DATE_OF_TERMINATION] >= '<FYStart>'))
ORDER BY OrgCode
