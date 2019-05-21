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
    PrintImageDimensions = [PSCustomObject]@{
        Width = 2605
        Height = 1051
    }
    ArtBoardDimensions = [PSCustomObject]@{
        Wdith = 2298
        Height = 811
    }
    OrderNumberPosition = @{
        Wdith = 85
        Hegith = 525
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
    ArtBoardDimensions = [PSCustomObject]@{
        Wdith = 2870
        Height = 757
    }
    OrderNumberPosition = @{
        Wdith = 50
        Hegith = 1100
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
        Width = 2764
        Height = 1640
    }
    ArtBoardDimensions = [PSCustomObject]@{
        Wdith = 2521
        Height = 1542
    }
    OrderNumberPosition = @{
        Wdith = 50
        Hegith = 1488
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
    OrderNumberPosition = @{
        Wdith = 50
        Hegith = 1696
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
    ArtBoardDimensions = [PSCustomObject]@{
        Width = 2961
        Height = 1094
    }
    OrderNumberPosition = @{
        Wdith = 150
        Hegith = 1062
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
    ArtBoardDimensions = [PSCustomObject]@{
        Width = 2625
        Height = 1620
    }
    OrderNumberPosition = @{
        Wdith = 1312
        Hegith = 810
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
    ArtBoardDimensions = [PSCustomObject]@{
        Width = 2965
        Height = 2219
    }
    OrderNumberPosition = @{
        Wdith = 50
        Hegith = 2165
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
    ArtBoardDimensions = [PSCustomObject]@{
        Width = 2679
        Height = 2279
    }
    OrderNumberPosition = @{
        Wdith = 50
        Hegith = 2181
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
        Width = 3395
        Height = 2409
    }
    ArtBoardDimensions = [PSCustomObject]@{
        Width = 3395
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
    ArtBoardDimensions = [PSCustomObject]@{
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
    ArtBoardDimensions = [PSCustomObject]@{
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

$TestCases = @(
    @{
        Size = 16
        FormType = "DWT"
        Description = "Augmented reality"
        Scene7Path = '/is/agm/tervis/16_cstm_print_mark?&setAttr.imgWrap={source=@Embed(%27is(tervisRender/16oz_wrap_final%3flayer=1%26src=ir(tervisRender/16_Warp_trans%3f%26obj=group%26decal%26src=is(tervisRender/16oz_base2%3f.BG%26layer=5%26anchor=0,0%26src=is(tervis/prj-8d47321e-dfea-4f98-ae86-8a57c85a78ad))%26show%26res=300%26req=object%26fmt=png-alpha,rgb)%26fmt=png-alpha,rgb)%27)}&setAttr.maskWrap={source=@Embed(%27http://images.tervis.com/is/image/tervis%3fsrc=(http://images.tervis.com/is/image/tervisRender/16oz_mark_mask%3f%26layer=1%26mask=is(tervisRender%3f%26src=ir(tervisRender/16_Warp_trans%3f%26obj=group%26decal%26src=is(tervisRender/16oz_base2%3f.BG%26layer=5%26anchor=0,0%26src=is(tervis/prj-8d47321e-dfea-4f98-ae86-8a57c85a78ad))%26show%26res=300%26req=object%26fmt=png-alpha)%26op_grow=-2)%26scl=1%26layer=2%26src=is(tervisRender/mark_mask_v1%3f%26layer=1%26mask=is(tervis/vum-8d47321e-dfea-4f98-ae86-8a57c85a78ad-66TU5O11)%26scl=1)%26scl=1)%26scl=1%26fmt=png8%26quantize=adaptive,off,2,ffffff,00A99C%27)}&setAttr.imgMark={source=@Embed(%27is(tervis/vum-8d47321e-dfea-4f98-ae86-8a57c85a78ad-66TU5O11)%27}&imageres=300&fmt=pdf,rgb&.v=3284}&$orderNum=11361062/2'
    },
    @{
        size = 6
        FormType = "SIP"
        Description = "Monogram high detail white ink areas"
        Scene7Path = '/is/agm/tervis/6_cstm_print?&setAttr.imgWrap={source=@Embed(%27is(tervisRender/6oz_wrap_final%3flayer=1%26src=ir(tervisRender/6_Warp_trans%3f%26obj=group%26decal%26src=is(tervisRender/6oz_base2%3f.BG%26layer=5%26anchor=0,0%26src=is(tervis/prj-aa1f3d62-dd31-411e-bc46-b7c963e77ae0))%26show%26res=300%26req=object%26fmt=png-alpha,rgb)%26fmt=png-alpha,rgb)%27)}&setAttr.maskWrap={source=@Embed(%27http://images.tervis.com/is/image/tervis%3fsrc=(http://images.tervis.com/is/image/tervisRender/6oz_wrap_mask%3f%26layer=1%26mask=is(tervisRender%3f%26src=ir(tervisRender/6_Warp_trans%3f%26obj=group%26decal%26src=is(tervisRender/6oz_base2%3f.BG%26layer=5%26anchor=0,0%26src=is(tervis/prj-aa1f3d62-dd31-411e-bc46-b7c963e77ae0))%26show%26res=300%26req=object%26fmt=png-alpha)%26op_grow=-2)%26scl=1)%26scl=1%26fmt=png8%26quantize=adaptive,off,2,ffffff,00A99C%27)}&imageres=300&fmt=pdf,rgb&.v=72271&$orderNum=11361062/2'
    },
    @{
        size = 20
        FormType = "SS"
        Description = "Stainless with high detail white ink areas"
        Scene7Path = '/is/agm/tervis/SS20_cstm_print?&setAttr.imgWrap=%7bsource=@Embed(%27is(tervisRender/SS_20oz_base2%3f.BG%26layer=5%26anchor=0,0%26src=is(tervis/prj-e5d6aaa1-f97e-4599-9a25-42d49b533409)%26fmt=png-alpha,rgb)%27)%7d&setAttr.maskWrap=%7bsource=@Embed(%27http://images.tervis.com/is/image/tervis%3fsrc=(http://images.tervis.com/is/image/tervisRender/SS_20oz_wrap_mask%3f%26layer=1%26mask=is(tervisRender/SS_20oz_base2%3f.BG%26layer=5%26anchor=0,0%26src=is(tervis/prj-e5d6aaa1-f97e-4599-9a25-42d49b533409)))%26op_grow=1%26op_usm=5,250,255,0%26scl=1%26cache=off%27)%7d&imageres=300&fmt=pdf,cmyk&cache=off'
    },
    @{
        
    }
)