/****** Object:  StoredProcedure [dbo].[usp_CML_Validate_RLO_Item_Step2]    Script Date: 04/01/2016 18:36:20 ******/
USE [DKSLS01]
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[usp_CML_Validate_RLO_Item_Step2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_CML_Validate_RLO_Item_Step2]
GO


CREATE PROCEDURE [dbo].[usp_CML_Validate_RLO_Item_Step2]

	@M_KD				Int,
	@ReturnReasonCode	Int = NUll,	
	@RLO_ReturnType		Varchar(32) OUTPUT ,
	@ADJ_Code			Int OUTPUT ,
	@PDT_Msg			Varchar(20) OUTPUT 	
AS
--  
--            
Declare @Return_Code				Int 
Declare @F_KD_RECALL				char(1)	
Declare @k_KeycodeNotFound			int 

Declare @k_RLO_ReturnType_Recall		Varchar(32)
Declare @k_RLO_ReturnType_Claimable		Varchar(32)
Declare @k_RLO_ReturnType_NonClaimable	Varchar(32)
Declare @k_RLO_ReturnType_WriteOff		Varchar(32)
Declare @k_Claimable					int 
Declare @k_WriteOff_Amount				Money

Declare @A_Chrg_Out_Cost				Money
Declare @Unt_Cst						Money
Declare @CML_Disp_CD					int
Declare @C_MDept						Char(3)
Declare @Flag_ClothingFootwearHeaterChristmas	Char(1)

Declare @RLO_ADJCode_Claimable_SIT		Int
Declare @RLO_ADJCode_WriteOff			Int
Declare @RLO_ADJCode_NonClaimable		Int

Set Nocount On

Set @Return_Code		= 1  
Set @ReturnReasonCode	= ''
Set @k_Claimable		= 7
Set @k_WriteOff_Amount	= 4.0

Set @k_KeycodenotFound				= 1
Set @k_RLO_ReturnType_Recall		= 'Recall'
Set @k_RLO_ReturnType_Claimable		= 'Claimable'
set @k_RLO_ReturnType_NonClaimable	= 'Salvage'
Set @k_RLO_ReturnType_WriteOff		= 'WriteOff'

Set @A_Chrg_Out_Cost				= 0.0
set @Unt_Cst						= 0.0
set @CML_Disp_CD					= 0 
Set @Flag_ClothingFootwearHeaterChristmas	= 'N'

set @RLO_ADJCode_Claimable_SIT		= 41
Set @RLO_ADJCode_WriteOff			= 11
Set @RLO_ADJCode_NonClaimable		= 3
Set @RLO_ReturnType					= ''

Set @ADJ_Code						= 0 
Set @PDT_Msg						= Space(20)


Print RTRIM(CONVERT(varchar(30), GETDATE()))  +  '  usp_Validate_RLO_Item_Step2   Started'

	Select @F_KD_RECALL = isnull(F_KD_RECALL,' ')
	       from dksls01.dbo.slt200 Where m_kd = @M_KD 
	
		if @@RowCount >  0 
			Begin
				if @F_KD_RECALL = 'Y'
					Begin
						Select @RLO_ReturnType = RLO_ReturnType
						from dksls01.dbo.RLO_RecallItems
						where m_kd = @M_KD
						and getdate() > d_eftve

						if @RLO_ReturnType = @k_RLO_ReturnType_NonClaimable
							Begin
								Set @Return_Code = 0
								Set @RLO_ReturnType = @k_RLO_ReturnType_NonClaimable
								Set @ADJ_Code = @RLO_ADJCode_NonClaimable
								Set @PDT_Msg = 'Salvage Pallet!'
								GOTO EndSave
							End

						if @RLO_ReturnType = @k_RLO_ReturnType_Claimable
							Begin
								Set @Return_Code = 0
								Set @RLO_ReturnType = @k_RLO_ReturnType_Claimable
								Set @ADJ_Code = @RLO_ADJCode_Claimable_SIT
								Set @PDT_Msg = 'Claimable Pallet!'
								GOTO EndSave
							End

						Set @Return_Code = 0
						Set @RLO_ReturnType = @k_RLO_ReturnType_Recall
						Set @PDT_Msg = 'Recall Item!'
						GOTO EndSave	
					End
			End


	Select Top 1 M_KD from dksls01.dbo.SLT200 Where M_KD = @M_KD and
			(	-- Footwear
				convert(int,C_Mdept) in (1,9,14,55,81)
				or  --  Clothing
				convert(int,C_Mdept) in (2,4,5,6,7,8,10,11,12,13,16,19,26,30,34,39,52,60,64,65,70,82,84,88,90,91,92,93,96,36)
				or -- Heater
				(C_Mdept = '056' and C_MCHL_Two = '600' and C_MCHL_Three = '004' and C_MCHL_Four = '001' and C_MCHL_Five = '002' and C_MCHL_Six = '001') 
				or -- Christmas and Easter
				convert(int,C_Mdept) in (17,21,29))

		if @@RowCount =  0 
				Set @Flag_ClothingFootwearHeaterChristmas = 'N'
			Else 
				Set @Flag_ClothingFootwearHeaterChristmas = 'Y'

	if @Flag_ClothingFootwearHeaterChristmas = 'Y'		
			Begin
				Set @Return_Code = 0
				Set @RLO_ReturnType = @k_RLO_ReturnType_Claimable
				Set @ADJ_Code = @RLO_ADJCode_Claimable_SIT
				Set @PDT_Msg = 'Claimable Pallet!'
			End
		Else
			Begin
				Set @Return_Code = 0
				Set @RLO_ReturnType = @k_RLO_ReturnType_NonClaimable
				Set @ADJ_Code = @RLO_ADJCode_NonClaimable
				Set @PDT_Msg = 'Salvage Pallet!'
			End
			
	Print RTRIM(CONVERT(varchar(30), GETDATE()))  +  '  usp_Validate_RLO_Item_Step2  Successfully'

GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0)
    BEGIN
        print 'ROLLBACK TRANSACTION'
    END
    return @Return_Code
EndSave:
return @Return_Code
GO

GRANT EXECUTE ON [dbo].[usp_CML_Validate_RLO_Item_Step2] TO [StoreRO]

GO
GRANT EXECUTE ON [dbo].[usp_CML_Validate_RLO_Item_Step2] TO [StoreRW]

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO
