/****** Object:  StoredProcedure [dbo].[usp_CML_Validate_RLO_Item_Step1]    Script Date: 04/01/2016 18:36:20 ******/
USE [DKSLS01]
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[usp_CML_Validate_RLO_Item_Step1]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_CML_Validate_RLO_Item_Step1]
GO

CREATE PROCEDURE [dbo].[usp_CML_Validate_RLO_Item_Step1]

	@M_KD				Int,
  	@ReturnReasonCode	Int = NULL,
  	@RLO_ReasonSubCode  Varchar(3),	
  	@RLO_ReturnType		Varchar(32) OUTPUT ,
  	@ADJ_Code			Int OUTPUT ,
  	@PDT_Msg			Varchar(20) OUTPUT,
  	@Unt_Cst            Money OUTPUT 	
AS
--  
--            
Declare @Return_Code				Int 
Declare @F_KD_RECALL				char(1)	
Declare @k_KeycodeNotFound			int 

Declare @k_RLO_ReturnType_Recall		           Varchar(32)
Declare @k_RLO_ReturnType_Claimable		           Varchar(32)
Declare @k_RLO_ReturnType_NonClaimable	           Varchar(32)
Declare @k_RLO_ReturnType_WriteOff		           Varchar(32)
Declare @k_RLO_ReturnType_PromptOnClearanceTrolley Varchar(32)
Declare @k_RLO_ReturnType_HoldForTesting           Varchar(32)
Declare @k_RLO_ReturnType_FoodItem                 Varchar(32)
Declare @k_RLO_ReturnType_PersonalItem             Varchar(32)

Declare @k_Claimable					int
Declare	@k_WriteOff_Amount				Money
Declare	@k_Claimable_Amount				Money
Declare @A_PR_POS						Money
Declare @Sell_Prc						Money
Declare @CML_Disp_CD					int
Declare @C_MDept						Char(3)
Declare @C_Mchl_Two                     Char(3)
Declare @C_Mchl_Three                   Char(3)
Declare @Food_Item                      Char(1)
Declare @Personal_Item                  Char(1)

Declare @RLO_ADJCode_Claimable_SIT		Int
Declare @RLO_ADJCode_WriteOff			Int
Declare @RLO_ADJCode_NonClaimable		Int

Set Nocount On

Set @Return_Code		= 1 
Set @ReturnReasonCode	= ''
Set @k_Claimable		= 7
Set @k_WriteOff_Amount	= 2.0
Set @k_Claimable_Amount	= 4.0
Set @Food_Item          = 'N'
Set @Personal_Item      = 'N'

Set @k_KeycodenotFound				           = 1
Set @k_RLO_ReturnType_Recall		           = 'Recall'
Set @k_RLO_ReturnType_Claimable		           = 'Claimable'
set @k_RLO_ReturnType_NonClaimable	           = 'Salvage'
set @k_RLO_ReturnType_PromptOnClearanceTrolley = 'PromptOnClearanceTrolley'
Set @k_RLO_ReturnType_WriteOff		           = 'WriteOff'
Set @k_RLO_ReturnType_HoldForTesting           = 'HoldForTesting'
Set @k_RLO_ReturnType_FoodItem                 = 'FoodItem'
Set @k_RLO_ReturnType_PersonalItem             = 'PersonalItem'

Set @A_PR_POS           			 = 0.0
set @Unt_Cst						 = 0.0
set @CML_Disp_CD					 = 0 

set @RLO_ADJCode_Claimable_SIT		 = 41
Set @RLO_ADJCode_WriteOff			 = 11
Set @RLO_ADJCode_NonClaimable		 = 3
Set @RLO_ReturnType					 = ''

Set @ADj_Code = 0
Set @PDT_Msg						 = Space(20)

Print RTRIM(CONVERT(varchar(30), GETDATE()))  +  '  usp_Validate_RLO_Item_Step1  Started'

	Select @F_KD_RECALL = isnull(F_KD_RECALL,' '), @A_PR_POS = IsNull(A_PR_POS, 0.0), 
	       @C_MDept = C_MDept, @C_Mchl_Two = C_Mchl_Two, @C_Mchl_Three = C_Mchl_Three 
	       from dksls01.dbo.slt200 Where m_kd = @M_KD
		if @@RowCount =  0 
			Begin
				Set @Return_Code = @k_KeycodeNotFound
				GOTO QuitWithRollback
			End
		else
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
				
                if @C_MDept = '021' Or @C_MDept = '029' Or @C_MDept = '044' Or @C_MDept = '045' Or (@C_MDept = '046' And @C_Mchl_Two = '600' And @C_Mchl_Three = '001') 
                    Begin
				        Set @Food_Item = 'Y'
				    End
				    
                if @C_MDept = '020' Or @C_MDept = '085'				
                    Begin
				        Set @Personal_Item = 'Y'
				    End
			End

	Select @CML_Disp_CD = isnull(CML_Disp_CD,0) , @Unt_Cst = isnull(Unt_Cst,0.0) from gstore.dbo.Item Where convert(int,Itm_ID) = @M_KD 
		if @@RowCount =  0 
			Begin
				Set @Return_Code = @k_KeycodeNotFound
				GOTO QuitWithRollback
			End
  		else
  			if @CML_Disp_CD = @k_Claimable And @Food_Item = 'N' And @Personal_Item = 'N' And @Unt_Cst >= @k_Claimable_Amount	-- Claimable Logic
  				Begin
  					Set @Return_Code = 0
  					Set @RLO_ReturnType = @k_RLO_ReturnType_Claimable
  					Set @ADj_Code = @RLO_ADJCode_Claimable_SIT	
  					Set @PDT_Msg = 'Claimable Item!'
  					--GOTO EndSave										
  				End
  			Else	-- Non Claimable
  				Begin
  				    Select M_KD from dksls01.dbo.RLO_HoldForTesting Where m_kd = @M_KD
  				
					if @@RowCount > 0 And @Food_Item = 'N' And @Personal_Item = 'N'
						Begin
		  					Set @Return_Code = 0
		  					Set @RLO_ReturnType = @k_RLO_ReturnType_HoldForTesting
		  					Set @ADj_Code = @RLO_ADJCode_NonClaimable	
		  					Set @PDT_Msg = 'Claimable Pallet!'
						End
  				    
  				    If @Return_Code <> 0

	  					if @Unt_Cst < @k_WriteOff_Amount							
	  						Begin		
	  							Set @Return_Code = 0
	  							Set @RLO_ReturnType = @k_RLO_ReturnType_WriteOff
	  							Set @ADj_Code = @RLO_ADJCode_WriteOff	
	  							Set @PDT_Msg = 'Throw Out Item!'
	  							--GOTO EndSave										
	  						End					
	        			else
	  						if (@RLO_ReasonSubCode = '020' or @RLO_ReasonSubCode = '025') And @Food_Item = 'N' And @Personal_Item = 'N'
	  							Begin
	  								Set @Return_Code = 0
	  								Set @RLO_ReturnType = @k_RLO_ReturnType_NonClaimable
	  								Set @PDT_Msg = 'Return Item ?'
	  								--GOTO EndSave										
	  							End
	  						Else
								if @Food_Item = 'Y'
		  							Begin
		  								Set @Return_Code = 0
		  								Set @RLO_ReturnType = @k_RLO_ReturnType_FoodItem
		  								Set @PDT_Msg = 'Food Item'
		  								--GOTO EndSave										
		  							End
								Else
									if @Personal_Item = 'Y'
			  							Begin
			  								Set @Return_Code = 0
			  								Set @RLO_ReturnType = @k_RLO_ReturnType_PersonalItem
			  								Set @PDT_Msg = 'Personal Item'
			  								--GOTO EndSave										
			  							End
								    Else
										Begin
											Set @Return_Code = 0
											Set @RLO_ReturnType = @k_RLO_ReturnType_PromptOnClearanceTrolley
											Set @PDT_Msg = 'RLO Item ?'
											--GOTO EndSave										
										End
  			    End				

Print RTRIM(CONVERT(varchar(30), GETDATE()))  +  '  usp_Validate_RLO_Item_Step1  Successfully'

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

GRANT EXECUTE ON [dbo].[usp_CML_Validate_RLO_Item_Step1] TO [StoreRO]

GO
GRANT EXECUTE ON [dbo].[usp_CML_Validate_RLO_Item_Step1] TO [StoreRW]

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO
