----------Updated-----------
NorthStar Database SP:

GetSeries nvarchar (30);
GetIndicator nvarchar (30);
GetIndicator1 nvarchar (30);
temp_var_0 integer;
temp_var_1 integer;

IF :object_type = '4' AND (:transaction_type = ('A')) THEN 

SELECT "Series" into GetSeries FROM OITM T0 where  T0."ItemCode" = :list_of_cols_val_tab_del; 
Select "SeriesName" into GetSeries from NNM1 where "Series"=:GetSeries;
IF (:GetSeries = 'Saudi') Then

Select T0."Indicator",T1."Indicator" into GetIndicator,GetIndicator1 from OFPR T0 Left join "DB_SRMTEST".OFPR T1 on T0."F_RefDate"=T1."F_RefDate"  and T0."T_RefDate"=T1."T_RefDate"
where CURRENT_DATE between T0."F_RefDate"  and T0."T_RefDate" ;
 
Select T0."NextNumber"-1, T1."NextNumber",T0."SeriesName" into temp_var_0, temp_var_1, GetSeries from NNM1 T0 Left join "DB_SRMTEST".NNM1 T1 on T0."ObjectCode"=T1."ObjectCode" and T0."SeriesName"=T1."SeriesName"
where T0."ObjectCode"='4' and T0."SeriesName"='Saudi' and (T0."Indicator"=:GetIndicator or T1."Indicator"=:GetIndicator1);

IF (:temp_var_0 > :temp_var_1) then
error := 010005;
error_message := ' Please Update the ' || temp_var_1 ||' ending Item. Since Item Creation is not possible in that DB_SRMTEST DB.';
End if;
End if;
End if;


Hilal Destination Database SP:

IF :object_type = '4' AND (:transaction_type = ('A') OR :transaction_type ='U') THEN 

SELECT (SELECT Count("DocEntry") FROM OITM 
        where  "ItemCode" = :list_of_cols_val_tab_del and ifnull("U_ItemSync",'') = 'N')
into temp_var_0  from Dummy;
if (:temp_var_0 > 0) then
error := 100025;
error_message := :list_of_cols_val_tab_del ||' Item is Not Created in NorthStar DB Please Create in appropriate 
DB...';
End if;
End if;

Common SP for preventing the item from removing: 

IF (:object_type = '4' AND :transaction_type = 'D') THEN 
if (1 = 1) then
error := 60002;
error_message := ' You are not authorized for deleting the Item...';
End if;
End if;
---------------------------------------------------
IF :object_type = '4' AND (:transaction_type = ('A') OR :transaction_type ='U') THEN 

SELECT (SELECT Count("DocEntry") FROM OITM 
        where  "ItemCode" = :list_of_cols_val_tab_del and ifnull("U_ItemSync",'') = 'N')
into temp_var_0  from Dummy;
if (:temp_var_0 > 0) then
error := 01;
error_message := :list_of_cols_val_tab_del ||' Item is Not Created in NorthStar DB Please Create in appropriate DB...';
End if;
End if;

IF :object_type = '4' AND (:transaction_type = ('A') OR :transaction_type ='U') THEN 

SELECT (SELECT Count("DocEntry") FROM OITM 
        where  "ItemCode" = :list_of_cols_val_tab_del and ifnull("U_ItemSync",'') = 'N')
into temp_var_0  from Dummy;
if (:temp_var_0 > 0) then
error := 01;
error_message := :list_of_cols_val_tab_del ||' Item is Not Created in NorthStar DB Please Create in appropriate DB...';
End if;
End if;

