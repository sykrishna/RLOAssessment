/****** Object:  StoredProcedure [dbo].[usp_CML_Validate_RLO_Item_Step3]    Script Date: 04/01/2016 18:36:20 ******/
USE [DKSLS01]
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[usp_CML_Validate_RLO_Item_Step3]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_CML_Validate_RLO_Item_Step3]
GO


CREATE PROCEDURE [dbo].[usp_CML_Validate_RLO_Item_Step3]

	@M_KD				Int,
	@sAPN				Varchar(13),
	@RLO_ReturnType		Varchar(32),
	@Item_ReturnType	Varchar(32),
	@Item_ReasonCode	Int,
	@SSCC_ID      		Varchar(32),
	@USR_ID      		Varchar(12),
	@PDT_Msg			Varchar(120) OUTPUT 	
AS
--  
--            
Declare @Return_Code				Int 
Declare @k_KeycodeNotFound			int 

Declare @k_RLO_ReturnType_Recall		Varchar(32)
Declare @k_RLO_ReturnType_Claimable		Varchar(32)
Declare @k_RLO_ReturnType_NonClaimable	Varchar(32)
Declare @k_RLO_ReturnType_WriteOff		Varchar(32)
Declare @k_Claimable					int 
Declare @CML_Disp_CD					int
Declare @ADJ_Code						int
Declare @SSCC_Count						int
Declare @Item_ReturnType_Code			int

Declare @Unt_Cost						Money
Declare @Sell_Price						Money
Declare @Dept							int
Declare @BatchID						Bigint

Declare @SSCC_ID_Existing	      		Varchar(32)
Declare @SSCC_ID_1st					Varchar(32)
Declare @SSCC_ID_2nd					Varchar(32)

Declare @RLO_ADJCode_Claimable_SIT		Int
Declare @RLO_ADJCode_WriteOff			Int
Declare @RLO_ADJCode_NonClaimable		Int

Declare @k_Return_Code_OK				Int
Declare @k_Return_Code_WrongClaimType	Int
Declare @k_Return_Code_SCM_AlreadyUsed  Int 
Declare @k_Return_Code_WrongSCM			Int 
Declare @k_BatchCompleted				Int

Declare	@M_APN							Bigint

Set Nocount On

Set @Return_Code		= 1  

Set @k_KeycodenotFound				= 1

Set @k_Return_Code_OK				= 0 
Set @k_Return_Code_WrongClaimType	= 1
Set @k_Return_Code_SCM_AlreadyUsed  = 2 
Set @k_Return_Code_WrongSCM			= 3
Set @k_BatchCompleted				= 12

Set @k_RLO_ReturnType_Recall		= 'Recall'
Set @k_RLO_ReturnType_Claimable		= 'Claimable'
set @k_RLO_ReturnType_NonClaimable	= 'Salvage'
Set @k_RLO_ReturnType_WriteOff		= 'WriteOff'

set @RLO_ADJCode_Claimable_SIT		= 41
Set @RLO_ADJCode_WriteOff			= 11
Set @RLO_ADJCode_NonClaimable		= 3

Set @ADJ_Code						= 0 
Set @PDT_Msg						= Space(20)

Set @SSCC_Count						= 0 
Set @Unt_cost						= 0.0
Set @Sell_Price						= 0.0
Set @Dept							= 0
Set @BatchID						= 0

Set @SSCC_ID_Existing				= ''
Set @SSCC_ID_1st					= ''
Set @SSCC_ID_2nd					= ''


Print RTRIM(CONVERT(varchar(30), GETDATE()))  +  '  usp_Validate_RLO_Item_Step3   Started'

Set @sAPN = isnull(@sAPN,'0')
Set @m_APN = convert(bigint,@sAPN)

	Select @Unt_cost = isnull(Unt_Cst,0) , @Sell_Price = isnull(RTL_PRC,0) from Gstore.dbo.Item Where ITM_ID = @m_KD
	Select @Dept = Cast(isnull(c_mdept,'0') As int) from DKSLS01.dbo.SLT200 Where M_KD = @m_KD	
	--
	--  RLO Claim Type must be either these two 
	--
	if (@RLO_ReturnType <> @k_RLO_ReturnType_Claimable and @RLO_ReturnType <> @k_RLO_ReturnType_NonClaimable)
				Begin
					Set @Return_Code = @k_Return_Code_WrongClaimType
					Set @PDT_Msg	 = left('Wrong Pallet Type'+space(20),20)
					GOTO EndSave
				End
	--
	--  Item Claim Type must be either these two 
	--
	Set @ADJ_Code = @RLO_ADJCode_NonClaimable
	if @Item_ReturnType = @k_RLO_ReturnType_Claimable 
				Begin 
					Set @Item_ReturnType_Code = 1 
					Set @ADJ_Code = @RLO_ADJCode_Claimable_SIT
			    End
			Else 
				if @Item_ReturnType = @k_RLO_ReturnType_NonClaimable
						Set @Item_ReturnType_Code = 3 
					Else
						Begin
							Set @Item_ReturnType_Code = 0 
							Set @Return_Code = @k_Return_Code_WrongClaimType
							Set @PDT_Msg	 = left('Wrong Claim Type'+space(20),20)
							GOTO EndSave
						End
	--
	--  Check SSCC_ID already used and Despatch
	--
	Select @SSCC_ID_Existing = SSCC_ID from dksls01.dbo.RLO_Claim_DTL 
		Where (isnull(RLO_ManifestNo,'') > '' and SSCC_ID = @SSCC_ID) 
		if @@RowCount > 0 
				Begin
					Set @Return_Code = @k_Return_Code_SCM_AlreadyUsed
					
                                        if @RLO_ReturnType = @k_RLO_ReturnType_Claimable
                                            Begin                                            
                                                Select @SSCC_ID_Existing = SSCC_ID from dksls01.dbo.RLO_Claim_DTL 
     		                                Where RLO_ManifestNo is NULL and RLO_ReturnType = @k_RLO_ReturnType_Claimable and Void_SCM is NULL

                                                if @@RowCount > 0
                                                    Set @PDT_Msg	 = left('SCM Already Sent'+space(20),20) + left(@SSCC_ID_Existing + space(20),20)
                                                Else
                                                    Set @PDT_Msg	 = left('SCM Already Sent'+space(20),20) + left(space(20),20)

                                            End
                                        Else
                                            Begin
                                                Select @SSCC_ID_Existing = SSCC_ID from dksls01.dbo.RLO_Claim_DTL 
     		                                Where RLO_ManifestNo is NULL and RLO_ReturnType = @k_RLO_ReturnType_NonClaimable and Void_SCM is NULL

                                                if @@RowCount > 0
                                                    Begin			                            
                                                        Select @SSCC_ID_1st = min(SSCC_ID) , @SSCC_ID_2nd = max(SSCC_ID) from DKsls01.dbo.RLO_Claim_DTL Where isnull(RLO_ManifestNo,'') = '' and RLO_ReturnType = @RLO_ReturnType and Void_SCM is NULL group by RLO_ReturnType

                                                        if @SSCC_ID_1st <> @SSCC_ID_2nd
                                                            Set @PDT_Msg	 = left('SCM Already Sent'+space(20),20) + left(@SSCC_ID_1st+space(20),20) + left(@SSCC_ID_2nd+space(20),20)
                                                        Else
                                                            Set @PDT_Msg	 = left('SCM Already Sent'+space(20),20) + left(@SSCC_ID_1st + space(20),20)
                                                    End
                                                Else
                                                    Set @PDT_Msg	 = left('SCM Already Sent'+space(20),20) + left(space(20),20)
                                            End
					GOTO EndSave
				End
	--
	--  Check SSCC_ID already used with other claim Type 
	--
	Select @SSCC_ID_Existing = SSCC_ID from dksls01.dbo.RLO_Claim_DTL 
		Where RLO_ReturnType <> @RLO_ReturnType and SSCC_ID = @SSCC_ID
		if @@RowCount > 0 
				Begin
					Set @Return_Code = @k_Return_Code_SCM_AlreadyUsed

                                        if @RLO_ReturnType = @k_RLO_ReturnType_Claimable
                                            Begin                                            
                                                Select @SSCC_ID_Existing = SSCC_ID from dksls01.dbo.RLO_Claim_DTL 
     		                                Where RLO_ManifestNo is NULL and RLO_ReturnType = @k_RLO_ReturnType_Claimable and Void_SCM is NULL

                                                if @@RowCount > 0
                                                    Set @PDT_Msg	 = left('Wrong SCM scanned !'+space(20),20) + left(@SSCC_ID_Existing + space(20),20)
                                                Else
                                                    Set @PDT_Msg	 = left('Wrong SCM scanned !'+space(20),20) + left(space(20),20)

                                            End
                                        Else
                                            Begin
                                                Select @SSCC_ID_Existing = SSCC_ID from dksls01.dbo.RLO_Claim_DTL 
     		                                Where RLO_ManifestNo is NULL and RLO_ReturnType = @k_RLO_ReturnType_NonClaimable and Void_SCM is NULL

                                                if @@RowCount > 0
                                                    Begin			                            
                                                        Select @SSCC_ID_1st = min(SSCC_ID) , @SSCC_ID_2nd = max(SSCC_ID) from DKsls01.dbo.RLO_Claim_DTL Where isnull(RLO_ManifestNo,'') = '' and RLO_ReturnType = @RLO_ReturnType and Void_SCM is NULL group by RLO_ReturnType

                                                        if @SSCC_ID_1st <> @SSCC_ID_2nd 
                                                            Set @PDT_Msg	 = left('Wrong SCM scanned !'+space(20),20) + left(@SSCC_ID_1st+space(20),20) + left(@SSCC_ID_2nd+space(20),20)
                                                        Else
                                                            Set @PDT_Msg	 = left('Wrong SCM scanned !'+space(20),20) + left(@SSCC_ID_1st + space(20),20)
                                                    End
                                                Else
                                                    Set @PDT_Msg	 = left('Wrong SCM scanned !'+space(20),20) + left(space(20),20)
                                            End
					GOTO EndSave
				End

	--
	--  Check SSCC_ID already used and has been voided 
	--
	Select @SSCC_ID_Existing = SSCC_ID from dksls01.dbo.RLO_Claim_DTL 
		Where RLO_ReturnType = @RLO_ReturnType and SSCC_ID = @SSCC_ID and Void_SCM is not NULL
		if @@RowCount > 0 
				Begin
					Set @Return_Code = @k_Return_Code_SCM_AlreadyUsed

                                        if @RLO_ReturnType = @k_RLO_ReturnType_Claimable
                                            Begin                                            
                                                Select @SSCC_ID_Existing = SSCC_ID from dksls01.dbo.RLO_Claim_DTL 
     		                                Where RLO_ManifestNo is NULL and RLO_ReturnType = @k_RLO_ReturnType_Claimable and Void_SCM is NULL

                                                if @@RowCount > 0
                                                    Set @PDT_Msg	 = left('Wrong SCM scanned !'+space(20),20) + left(@SSCC_ID_Existing + space(20),20)
                                                Else
                                                    Set @PDT_Msg	 = left('Wrong SCM scanned !'+space(20),20) + left(space(20),20)

                                            End
                                        Else
                                            Begin
                                                Select @SSCC_ID_Existing = SSCC_ID from dksls01.dbo.RLO_Claim_DTL 
     		                                Where RLO_ManifestNo is NULL and RLO_ReturnType = @k_RLO_ReturnType_NonClaimable and Void_SCM is NULL

                                                if @@RowCount > 0
                                                    Begin			                            
                                                        Select @SSCC_ID_1st = min(SSCC_ID) , @SSCC_ID_2nd = max(SSCC_ID) from DKsls01.dbo.RLO_Claim_DTL Where isnull(RLO_ManifestNo,'') = '' and RLO_ReturnType = @RLO_ReturnType and Void_SCM is NULL group by RLO_ReturnType

                                                        if @SSCC_ID_1st <> @SSCC_ID_2nd 
                                                            Set @PDT_Msg	 = left('Wrong SCM scanned !'+space(20),20) + left(@SSCC_ID_1st+space(20),20) + left(@SSCC_ID_2nd+space(20),20)
                                                        Else
                                                            Set @PDT_Msg	 = left('Wrong SCM scanned !'+space(20),20) + left(@SSCC_ID_1st + space(20),20)
                                                    End
                                                Else
                                                    Set @PDT_Msg	 = left('Wrong SCM scanned !'+space(20),20) + left(space(20),20)
                                            End
					GOTO EndSave
				End

	--
	--  Claimable Pallet Logic
	--	
	if @RLO_ReturnType = @k_RLO_ReturnType_Claimable
		Begin 
			--
			--  Check any open SSCC_ID for the same Claim Type 
			--
			Select top 1 @SSCC_ID_Existing = SSCC_ID from DKsls01.dbo.RLO_Claim_DTL 
				Where isnull(RLO_ManifestNo,'') = '' and RLO_ReturnType = @RLO_ReturnType and Void_SCM is NULL
				if @@RowCount > 0   
						Begin 
							if @SSCC_ID = @SSCC_ID_Existing
									Begin										
										if @ADJ_Code = @RLO_ADJCode_NonClaimable
												Begin
													Insert into DKSLS01.dbo.CML_EmptyPacket_HDR 
    												(Usr_ID, Creation_TimeStamp, Commit_TimeStamp, Stat)
    												Values
    												(@Usr_ID, GETDATE(), GETDATE(), @k_BatchCompleted)
    												Select @BatchID = @@Identity
    												
    												Insert Into DKSLS01.dbo.CML_EmptyPacket_DTL 
												    (Batch_no, Seq_No, Capture_TimeStamp, Keycode, APN, Capture_QTY, Reason_Code, Dept_No, Sell_Prc, Cost_Prc, SOH, StoreCount) 
												    Values
												    (@BatchID, 1, GETDATE(), @M_KD, @M_APN, 1, '03', @Dept, @Sell_Price, @Unt_cost, 0, 0)
												End

										insert into dksls01.dbo.RLO_Claim_DTL
											(m_store, M_KD, M_APN, QTY, Unt_Cost, Sell_Price, RLO_ReturnType, Item_ReturnType, Item_Return_ReasonCode, ADJ_Code, SSCC_ID, Usr_ID, Insert_DateTime)
											Values( Substring(@@Servername,3,4),@M_KD,@M_APN,1,@Unt_cost, @Sell_Price,@RLO_ReturnType,@Item_ReturnType_Code,@Item_ReasonCode,@ADJ_Code,@SSCC_ID,@Usr_ID, GETDATE())										
										Set @Return_Code = @k_Return_Code_OK
										Set @PDT_Msg	 = 'Item has been Saved !'
										GOTO EndSave
									End
							   else
									Begin										
										Set @Return_Code = @k_Return_Code_WrongSCM										
										Set @PDT_Msg	 = left('Wrong SCM scanned !'+space(20),20) + left(@SSCC_ID_Existing+space(20),20)
										GOTO EndSave
									End
						End
					Else
						--
						--  First Item in the SCM
						--
						Begin
								if @ADJ_Code = @RLO_ADJCode_NonClaimable
										Begin
											Insert into DKSLS01.dbo.CML_EmptyPacket_HDR 
											(Usr_ID, Creation_TimeStamp, Commit_TimeStamp, Stat)
											Values
											(@Usr_ID, GETDATE(), GETDATE(), @k_BatchCompleted)
											Select @BatchID = @@Identity
											
											Insert Into DKSLS01.dbo.CML_EmptyPacket_DTL 
										    (Batch_no, Seq_No, Capture_TimeStamp, Keycode, APN, Capture_QTY, Reason_Code, Dept_No, Sell_Prc, Cost_Prc, SOH, StoreCount) 
										    Values
										    (@BatchID, 1, GETDATE(), @M_KD, @M_APN, 1, '03', @Dept, @Sell_Price, @Unt_cost, 0, 0)
										End
								
								--insert into logic
								insert into dksls01.dbo.RLO_Claim_DTL
									(m_store, M_KD, M_APN, QTY, Unt_Cost, Sell_Price, RLO_ReturnType, Item_ReturnType, Item_Return_ReasonCode, ADJ_Code, SSCC_ID, Usr_ID, Insert_DateTime)
									Values( Substring(@@Servername,3,4),@M_KD,@m_APN,1,@Unt_cost, @Sell_Price,@RLO_ReturnType,@Item_ReturnType_Code,@Item_ReasonCode,@ADJ_Code,@SSCC_ID,@Usr_ID, GETDATE())										
								Set @Return_Code = @k_Return_Code_OK
								Set @PDT_Msg	 =  'Item has been Saved !'
								GOTO EndSave
						End
		End		
	--
	--  Salvage Pallet Logic
	--	
	if @RLO_ReturnType = @k_RLO_ReturnType_NonClaimable
		Begin 
			--
			--  Check any open SSCC_ID for the same Claim Type 
			--
			Select @SSCC_ID_1st = min(SSCC_ID) , @SSCC_ID_2nd = max(SSCC_ID) from DKsls01.dbo.RLO_Claim_DTL Where isnull(RLO_ManifestNo,'') = '' and RLO_ReturnType = @RLO_ReturnType and Void_SCM is NULL group by RLO_ReturnType
			Select @SSCC_Count = count(*) from 
				   (Select SSCC_ID from DKsls01.dbo.RLO_Claim_DTL Where isnull(RLO_ManifestNo,'') = '' and RLO_ReturnType = @RLO_ReturnType and Void_SCM is NULL group by SSCC_ID) b				
				   Where b.SSCC_ID is not null
				if @SSCC_Count <= 1   
						--  Allow Max Two Salvage Pallets
						Begin 
								if @ADJ_Code = @RLO_ADJCode_NonClaimable
										Begin
											Insert into DKSLS01.dbo.CML_EmptyPacket_HDR 
											(Usr_ID, Creation_TimeStamp, Commit_TimeStamp, Stat)
											Values
											(@Usr_ID, GETDATE(), GETDATE(), @k_BatchCompleted)
											Select @BatchID = @@Identity
											
											Insert Into DKSLS01.dbo.CML_EmptyPacket_DTL 
										    (Batch_no, Seq_No, Capture_TimeStamp, Keycode, APN, Capture_QTY, Reason_Code, Dept_No, Sell_Prc, Cost_Prc, SOH, StoreCount) 
										    Values
										    (@BatchID, 1, GETDATE(), @M_KD, @M_APN, 1, '03', @Dept, @Sell_Price, @Unt_cost, 0, 0)
										End

								insert into dksls01.dbo.RLO_Claim_DTL
											(m_store, M_KD, M_APN, QTY, Unt_Cost, Sell_Price, RLO_ReturnType, Item_ReturnType, Item_Return_ReasonCode, ADJ_Code, SSCC_ID, Usr_ID, Insert_DateTime)
											Values( Substring(@@Servername,3,4),@M_KD,@m_APN,1,@Unt_cost, @Sell_Price,@RLO_ReturnType,@Item_ReturnType_Code,@Item_ReasonCode,@ADJ_Code,@SSCC_ID,@Usr_ID, GETDATE())										

								Set @Return_Code = @k_Return_Code_OK
								Set @PDT_Msg	 = 'Item has been Saved !'
								GOTO EndSave
						End
				   else
						Begin
							Select top 1 @SSCC_ID_Existing = SSCC_ID from dKsls01.dbo.RLO_Claim_DTL 
							Where isnull(RLO_ManifestNo,'') = '' and RLO_ReturnType = @RLO_ReturnType and Void_SCM is NULL
								  and SSCC_ID = @SSCC_ID
									if @@RowCount > 0   
											Begin
												if @ADJ_Code = @RLO_ADJCode_NonClaimable
														Begin
															Insert into DKSLS01.dbo.CML_EmptyPacket_HDR 
		    												(Usr_ID, Creation_TimeStamp, Commit_TimeStamp, Stat)
		    												Values
		    												(@Usr_ID, GETDATE(), GETDATE(), @k_BatchCompleted)
		    												Select @BatchID = @@Identity
		    												
		    												Insert Into DKSLS01.dbo.CML_EmptyPacket_DTL 
														    (Batch_no, Seq_No, Capture_TimeStamp, Keycode, APN, Capture_QTY, Reason_Code, Dept_No, Sell_Prc, Cost_Prc, SOH, StoreCount) 
														    Values
														    (@BatchID, 1, GETDATE(), @M_KD, @M_APN, 1, '03', @Dept, @Sell_Price, @Unt_cost, 0, 0)
														End

												insert into dksls01.dbo.RLO_Claim_DTL
													(m_store, M_KD, M_APN, QTY, Unt_Cost, Sell_Price, RLO_ReturnType, Item_ReturnType, Item_Return_ReasonCode, ADJ_Code, SSCC_ID, Usr_ID, Insert_DateTime)
													Values( Substring(@@Servername,3,4),@M_KD,@m_APN,1,@Unt_cost, @Sell_Price,@RLO_ReturnType,@Item_ReturnType_Code,@Item_ReasonCode,@ADJ_Code,@SSCC_ID,@Usr_ID, GETDATE())										
												Set @Return_Code = @k_Return_Code_OK
												Set @PDT_Msg	 = 'Item has been Saved !'
												GOTO EndSave
											End
									   else
											Begin										
												Set @Return_Code = @k_Return_Code_WrongSCM																						
												Set @PDT_Msg	 = left('Wrong SCM scanned !'+space(20),20) + left(@SSCC_ID_1st+space(20),20) + left(@SSCC_ID_2nd+space(20),20)
												GOTO EndSave
											End
						End
		End	

			
	Print RTRIM(CONVERT(varchar(30), GETDATE()))  +  '  usp_Validate_RLO_Item_Step3  Successfully'

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

GRANT EXECUTE ON [dbo].[usp_CML_Validate_RLO_Item_Step3] TO [StoreRO]

GO
GRANT EXECUTE ON [dbo].[usp_CML_Validate_RLO_Item_Step3] TO [StoreRW]

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO
