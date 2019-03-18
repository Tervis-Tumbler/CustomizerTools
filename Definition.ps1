$CustomyzerEnvionrments = [PSCustomObject]@{
    Name = "Production"
    PackListFilesPathRoot = "\\mizerftp.tervis.com\MizerFTP\PRD-MizerFTP"
    PackListXMLDestinationPath = "\\sgs-tervis-mt\User Mailboxes\SGS-NOKOMIS\To Send"
    RequisitionDestinationPath = "/ebsapps/PRD/biztalk/Mizer/inbound"
    CustomyzerDatabasePasswordStatePasswordGUID = "a59786e2-8468-499f-9881-e44ed3a286b1"
    EmailAddressToRecieveXLSXPasswordStatePasswordGUID = "31c6cd88-fbf7-4547-b2a5-3ed9180bd059"
    PasswordStateAPIKeyPasswordGUID = "6ab983d1-49c5-42ca-80b0-da6aa79b60ea"
    FileShareAccountPasswordStateGUID = "10d7227f-5f0b-45f1-9d53-a9279fdcbf74"
    ScheduledTaskRepetitionIntervalName = "EverWorkdayAt1PM"
},
[PSCustomObject]@{
    Name = "Epsilon"
    PackListFilesPathRoot = "\\mizerftp.tervis.com\MizerFTP\EPS-MizerFTP"
    PackListXMLDestinationPath = "\\tervis.prv\applications\Customizer\SGS\Epsilon"
    RequisitionDestinationPath = "/ebsapps/SIT/biztalk/Mizer/inbound"
    CustomyzerDatabasePasswordStatePasswordGUID = "9e22d5cf-96d1-4efb-8d76-57fe1f8720d4"
    EmailAddressToRecieveXLSXPasswordStatePasswordGUID = "b4e466c2-5648-4ca6-ad73-b93e38e1af1f"
    PasswordStateAPIKeyPasswordGUID = "8cc4a5e5-62f0-4d73-b268-c4a03ef477a4"
    FileShareAccountPasswordStateGUID = "d5db01dc-c120-4f2e-bea4-5b7505c5db8a"
    ScheduledTaskRepetitionIntervalName = "EverWorkdayAt2PM"
},
[PSCustomObject]@{
    Name = "Delta"
    PackListFilesPathRoot = "\\mizerftp.tervis.com\MizerFTP\DLT-MizerFTP"
    PackListXMLDestinationPath = "\\tervis.prv\applications\Customizer\SGS\Delta"
    RequisitionDestinationPath = "/ebsapps/DEV/biztalk/Mizer/inbound"
    CustomyzerDatabasePasswordStatePasswordGUID = "47d1f398-38e1-4f44-8598-4da0a94be31a"
    EmailAddressToRecieveXLSXPasswordStatePasswordGUID = "dc426f75-b719-43fc-b004-3d02cbbd188a"
    PasswordStateAPIKeyPasswordGUID = "786643ab-87a0-414a-97ff-9942c0bd904d"
    FileShareAccountPasswordStateGUID = "4b8f6d82-82e1-4f21-a9ef-f7e83af4efc2"
    ScheduledTaskRepetitionIntervalName = "EverWorkdayAt2PM"
}

$Script:SizeAndFormTypeToImageTemplateNames = [PSCustomObject]@{
    Size = 6
    FormType = "SIP"
    # ArtBoardDimensions = [PSCustomObject]@{
    #     Width = 
    #     Height = 
    # }
    PrintImageDimensions = [PSCustomObject]@{
        Width = 2605
        Height = 1051
    }
    ImageTemplateName = [PSCustomObject]@{
        Final = "6oz_wrap_final"
        Mask = "6oz_wrap_mask"
        Vignette = "6_Warp_trans"
        Base = "6oz_base2"
        Print = @{
            Illustrator = "6_cstm_print.ai"
            InDesign = "6SIP-cstm-print.idml"
            Scene7 = "6_cstm_print"
        }
    }
},
[PSCustomObject]@{
    Size = 9
    FormType = "SWG","WINE"
    PrintImageDimensions = [PSCustomObject]@{
        Width = 3260
        Height = 962
    }
    ImageTemplateName = [PSCustomObject]@{
        Final = "9oz_wrap_final"
        Mask = "9oz_wrap_mask"
        Vignette = "9_Warp_trans"
        Base = "9oz_base2"
        Print = @{
            Illustrator = "9_cstm_print.ai"
            InDesign = "9WINE-cstm-print.idml"
            Scene7 = "9_cstm_print"
        }
    }
},
[PSCustomObject]@{
    Size = 10
    FormType = "WAV","DWT"
    PrintImageDimensions = [PSCustomObject]@{
        Width = 2767
        Height = 1640
    }
    ImageTemplateName = [PSCustomObject]@{
        Final = "10oz_wrap_final"
        Mask = "10oz_wrap_mask"
        Vignette = "10_Warp_trans"
        Base = "10oz_base2"
        Print = @{
            Illustrator = "10_cstm_print.ai"
            InDesign = "10WAV-cstm-print.idml"
            Scene7 = "10_cstm_print"
        }
    }
},
[PSCustomObject]@{
    Size = 16
    FormType = "DWT"
    ArtBoardDimensions = [PSCustomObject]@{
        Width = 2717
        Height = 1750
    }
    PrintImageDimensions = [PSCustomObject]@{
        Width = 3084
        Height = 1873
    }
    ImageTemplateName = [PSCustomObject]@{
        Final = "16oz_wrap_final"
        Mask = "16oz_wrap_mask"
        Vignette = "16_Warp_trans"
        Base = "16oz_base2"
        FinalWithERPNumber = "16oz_final_x2_v1"
        Print = @{
            Illustrator = "16_cstm_print.ai"
            InDesign = "16DWT-cstm-print.idml"
            Scene7 = "16_cstm_print"
            Scene7AR = "16_cstm_print_mark"
        }
        WhiteInkMask = "16oz_wrap_mask_black"
    }
},
[PSCustomObject]@{
    Size = 16
    FormType = "MUG"
    PrintImageDimensions = [PSCustomObject]@{
        Width = 3628
        Height = 1339
    }
    ImageTemplateName = [PSCustomObject]@{
        Final = "MUG_wrap_final"
        Mask = "MUG_wrap_mask"
        Vignette = "MUG_Warp_trans"
        Base = "MUG_base2"
        Print = @{
            Illustrator = "MUG_cstm_print.ai"
            InDesign = "16MUG-cstm-print.idml"
            Scene7 = "MUG_cstm_print"
        }
    }
},
[PSCustomObject]@{
    Size = 16
    FormType = "BEER"
    PrintImageDimensions = [PSCustomObject]@{
        Width = 3074
        Height = 1748
    }
    ImageTemplateName = [PSCustomObject]@{
        Final = "BEER_wrap_final"
        Mask = "BEER_wrap_mask"
        Vignette = "BEER_Warp_trans"
        Base = "BEER_base2"
        Print = @{
            Illustrator = "BEER_cstm_print.ai"
            InDesign = "16BEER-cstm-print.idml"
            Scene7 = "BEER_cstm_print"
        }
    }
},
[PSCustomObject]@{
    Size = 24
    FormType = "DWT"
    PrintImageDimensions = [PSCustomObject]@{
        Width = 3574
        Height = 2402
    }
    ImageTemplateName = [PSCustomObject]@{
        Final = "24oz_wrap_final"
        Mask = "24oz_wrap_mask"
        Vignette = "24_Warp_trans"
        Base = "24oz_base2"
        Print = @{
            Illustrator = "24_cstm_print.ai"
            InDesign = "24DWT-cstm-print.idml"
            Scene7 = "24_cstm_print"
        }
    }
},
[PSCustomObject]@{
    Size = 24
    FormType = "WB"
    PrintImageDimensions = [PSCustomObject]@{
        Width = 2977
        Height = 2420
    }
    ImageTemplateName = [PSCustomObject]@{
        Final = "WB_wrap_final"
        Mask = "WB_wrap_mask"
        Vignette = "WB_Warp_trans"
        Base = "WB_base2"
        Print = @{
            Illustrator = "WB_cstm_print.ai"
            InDesign = "24WB-cstm-print.idml"
            Scene7 = "WB_cstm_print"
        }
    }
}   ,
[PSCustomObject]@{
    Size = 30
    FormType = "SS"
    PrintImageDimensions = [PSCustomObject]@{
        Width = 3394
        Height = 2409
    }
    ImageTemplateName = [PSCustomObject]@{
       Final = ""
       Mask = "SS_30oz_wrap_mask"
       Vignette = ""
       Base = "SS_30oz_base2"
       Print = @{
            Illustrator = "SS30_cstm_print.ai"
            InDesign = "30SS-cstm-print.idml"
            Scene7 = "SS30_cstm_print"
        }
    }
},
[PSCustomObject]@{
    Size = 20
    FormType = "SS"
    PrintImageDimensions = [PSCustomObject]@{
        Width = 2974
        Height = 2032
    }
    ImageTemplateName = [PSCustomObject]@{
       Final = ""
       Mask = "SS_20oz_wrap_mask"
       Vignette = ""
       Base = "SS_20oz_base2"
       Print = @{
            Illustrator = "SS20_cstm_print.ai"
            InDesign = "20SS-cstm-print.idml"
            Scene7 = "SS20_cstm_print"
        }
    }
},
[PSCustomObject]@{
    Size = 24
    FormType = "SS"
    PrintImageDimensions = [PSCustomObject]@{
        Width = 2916
        Height = 2367
    }
    ImageTemplateName = [PSCustomObject]@{
       Final = ""
       Mask = "SS_24oz_wrap_mask"
       Vignette = ""
       Base = "SS_24oz_base2"
       Print = @{
            Illustrator = "SS24_cstm_print.ai"
            InDesign = "24SS-cstm-print.idml"
            Scene7 = "SS24_cstm_print"
        }
    }
}