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

$Script:SizeAndFormTypeToImageTemplateNames = Invoke-RestMethod -Uri https://unpkg.com/@tervis/tervisproductmetadata@1.0.13/TervisProductMetadata.json

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