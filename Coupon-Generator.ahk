#SingleInstance Force
#Warn All, StdOut
#Requires AutoHotkey v2.0

#Include <Bruno-Functions\ImportAllList>
#Include <GithubReleases\GithubReleases>

global version := "1.0.1"

global sim_temp     := "c:\SIM\TMP\"
global install_path := NewDir(A_AppData "\TheBrunoCA\Coupons\")
global daily_folder := NewDir(install_path A_Year "\" A_Mon "\" A_MDay)
global cli_sale_log := daily_folder "\cli_sale_log.txt"
global cli_rec_log  := daily_folder "\cli_rec_log.txt"

global printer := PrinterHelper()

global first_run    := true

global coupon_ini_path  := A_Temp "\" A_ScriptName "_Coupon.ini"
global coupon_xls_path  := A_Temp "\" A_ScriptName "_Coupon.xls"
FileInstall("c:\Users\bruno\OneDrive\Documentos\Repos\Coupon-Generator\Coupon.ini", coupon_ini_path, true)
FileInstall("c:\Users\bruno\OneDrive\Documentos\Repos\Coupon-Generator\Coupon.xls", coupon_xls_path, true)

A_TrayMenu.Add("Checar por atualizações", CheckUpdates)
A_TrayMenu.Add("Imprimir", PrintAway)
A_TrayMenu.Add("Configurações", ConfigGui)


PrintAway(args*){
    local c := Coupon()
    c.value := config_ini["config", "coupon_value", 50]
    SetTimer(PrintNewCoupons, 0)
    PrintGui(c)
    return
}

global config_ini := Ini(install_path "config.ini")
global coupons_ini := Ini(coupon_ini_path)
global ignored_tp := install_path "ignored_tp.txt"
GetIgnoredTp(){
    if FileExist(ignored_tp)
        return FileRead(ignored_tp, "UTF-8")
    FileAppend("CREDIARIO`n", ignored_tp, "UTF-8")
    return FileRead(ignored_tp, "UTF-8")
}
SetIgnoredTp(new){
    try FileDelete(ignored_tp)
    FileAppend(new, ignored_tp, "UTF-8")
}

SetTimer(PrintNewCoupons, config_ini["config", "timer_interval", 1000])

PrintNewCoupons(){
    global first_run
    local fr := first_run
    local coupon := GetNewCoupon(sim_temp "\" A_Year A_Mon "\" A_MDay "\ImpressaoVenda*", cli_sale_log)
    if coupon != ""{
        SetTimer(PrintNewCoupons, 0)
        PrintGui(coupon)
        return
    }
    first_run := fr
    coupon := GetNewCoupon(sim_temp "\" A_Year A_Mon "\" A_MDay "\Recebimento*", cli_rec_log)
    if coupon != ""{
        SetTimer(PrintNewCoupons, 0)
        PrintGui(coupon)
        return
    }
}

github := GithubReleases("TheBrunoCA", "Coupon-Generator")
CheckUpdates(args*){
    github.GetInfo()
    if github.IsUpToDate(version){
        MsgBox("O aplicativo está atualizado.")
        return
    }
    answer := MsgBox("Atualização " github.GetLatestReleaseVersion() " encontrada`nDeseja atualizar?", , "0x4")
    if answer == "Yes"
        github.Update(A_ScriptDir)
}


ConfigGui(args*){
    SetTimer(PrintNewCoupons, 0)
    local ConfigGui := Gui(, A_ScriptName " v" version)
    ConfigGui.OnEvent("Close", _Close)
    _Close(args*){
        ConfigGui.Destroy()
        SetTimer(PrintNewCoupons, config_ini["config", "timer_interval", 1000])
        return
    }
    ConfigGui.SetFont("S14")
    ConfigGui.AddText("xm", "Configurações")
    ConfigGui.SetFont("S8")
    ConfigGui.AddText("xm", "Intervalo de checagem: ")
    check_interval_edit := ConfigGui.AddEdit("Number yp w30", (config_ini["config", "timer_interval", 1000] / 1000))
    ConfigGui.AddText("xm", "Valor do cupom: ")
    coupon_value_edit := ConfigGui.AddEdit("yp w30 x+45", config_ini["config", "coupon_value", 50])
    ConfigGui.AddText("xm", "Tipos de pagamento para ignorar`nSeparado por linhas.")
    ignored_tp_edit := ConfigGui.AddEdit("xm w150 Multi R5", GetIgnoredTp())
    ask_no_matter_value_ckb := ConfigGui.AddCheckbox(, "Mostrar cupom mesmo se o valor não for suficiente.")
    ask_no_matter_value_ckb.Value := config_ini["config", "ask_no_matter_value", false]
    ConfigGui.AddText("xm", "Impressora: ")
    pt := ""
    for i, p in printer.printers{
        if p == config_ini["config", "printer", printer.default]
            pt := "Choose" i
    }
    printer_ddl := ConfigGui.AddDropDownList(pt " w150", printer.printers)
    ConfigGui.AddText("ys+20", "Nome do sorteio: ")
    ConfigGui.AddText(, "Descrição curta dos premios: ")
    ConfigGui.AddText(, "Data do sorteio: ")
    name_sort_edit := ConfigGui.AddEdit("ys+20 w150", config_ini["config", "sort_name", "Sorteio"])
    prize_name_edit := ConfigGui.AddEdit("w150", config_ini["config", "prize_name", "Diversos prêmios"])
    prize_date_edit := ConfigGui.AddEdit("w150", config_ini["config", "prize_date", "xx/xx/xxxx"])
    cancel_btn := ConfigGui.AddButton("xm", "Cancelar")
    cancel_btn.OnEvent("Click", _Close)
    submit_btn := ConfigGui.AddButton("yp Default", "Salvar")
    submit_btn.OnEvent("Click", _Submit)
    _Submit(args*){
        global config_ini
        if coupon_value_edit.Value == ""
            return
        if not IsNumber(coupon_value_edit.Value)
            return
        if coupon_value_edit.Value < 1
            coupon_value_edit.Value := 1

        if check_interval_edit.Value == ""
            return
        if check_interval_edit.Value < 1
            check_interval_edit.Value := 1
        
        config_ini["config", "timer_interval"] := (check_interval_edit.Value * 1000)
        config_ini["config", "coupon_value"] := coupon_value_edit.Value
        SetIgnoredTp(ignored_tp_edit.Value)
        config_ini["config", "ask_no_matter_value"] := ask_no_matter_value_ckb.Value
        config_ini["config", "printer"] := printer_ddl.Text
        config_ini["config", "sort_name"] := name_sort_edit.Value
        config_ini["config", "prize_name"] := prize_name_edit.Value
        config_ini["config", "prize_date"] := prize_date_edit.Value
        ConfigGui.Destroy()
        SetTimer(PrintNewCoupons, config_ini["config", "timer_interval", 1000])
        return
    }
    ConfigGui.Show()
}

PrintGui(coupon){
    loop read ignored_tp{
        if InStr(coupon.payment_type, A_LoopReadLine) and coupon.payment_type != ""{
            _Cancel()
            return
        }
    }

    local PrintGui := Gui()
    PrintGui    .OnEvent("Close", _Cancel)
    PrintGui    .AddText(, "Nome:")
    cn := coupon.name == "CONSUMIDOR BALCAO" ? "" : coupon.name
    name_edit   := PrintGui.AddEdit("Multi R2 w200", cn)
    PrintGui    .AddText()
    PrintGui    .AddText(, "Número de Telefone:")
    tel_edit    := PrintGui.AddEdit("Multi R2 w200", coupon.tel)
    PrintGui    .AddText()
    PrintGui    .AddText(, "Cidade:")
    city_edit   := PrintGui.AddEdit("Multi R2 w200", coupon.city)
    PrintGui    .AddText()
    PrintGui    .AddText(, "Bairro:")
    dist_edit   := PrintGui.AddEdit("Multi R2 w200", coupon.dist)
    PrintGui    .AddText()
    PrintGui    .AddText(, "Endereço:")
    end_edit    := PrintGui.AddEdit("Multi R2 w200", coupon.end)
    PrintGui    .AddText()
    PrintGui    .AddText(, "Valor total: R$" coupon.value)
    cv := config_ini["config", "coupon_value", 50]
    if not InStr(cv, ",") or InStr(cv, ".")
        cv .= ",00"
    PrintGui    .AddText(, "Valor por cupom: R$" cv)
    cqtd := Floor(coupon.value / config_ini["config", "coupon_value", 50])
    if cqtd < 1 and config_ini["config", "ask_no_matter_value", false] == false{
        _Cancel()
        return
    }
        

    PrintGui    .AddText(, "Qtde correta de cupons: " cqtd)
    
    cancel_btn  := PrintGui.AddButton(,"Cancelar")
    cancel_btn  .OnEvent("Click", _Cancel)
    _Cancel(args*){
        PrintGui.Destroy()
        SetTimer(PrintNewCoupons, config_ini["config", "timer_interval", 1000])
    }
    
    submit_btn  := PrintGui.AddButton("Default yp", "Confirmar")
    submit_btn  .OnEvent("Click", _Print)
    _Print(args*){
        PrintGui.Hide()

        Sleep(100)
        obj := ComObject("Excel.Application")
        obj.WorkBooks.Open(coupon_xls_path, , false, , , , , , , , false)

        obj.ActiveSheet.Range(coupons_ini["positions", "prize_draw_name"]).Value := config_ini["config", "sort_name", "Sorteio"]
        obj.ActiveSheet.Range(coupons_ini["positions", "prizes"]).Value := config_ini["config", "prize_name", "Vários Prêmios"]
        obj.ActiveSheet.Range(coupons_ini["positions", "date"]).Value := config_ini["config", "prize_date", ""]
        obj.ActiveSheet.Range(coupons_ini["positions", "name"]).Value := "Nome: " name_edit.Value
        obj.ActiveSheet.Range(coupons_ini["positions", "tel"]).Value := "Telefone: " tel_edit.Value
        obj.ActiveSheet.Range(coupons_ini["positions", "city"]).Value := "Cidade: " city_edit.Value
        obj.ActiveSheet.Range(coupons_ini["positions", "dist"]).Value := "Bairro: " dist_edit.Value
        obj.ActiveSheet.Range(coupons_ini["positions", "end"]).Value := "Endereço: " end_edit.Value
        
        if copies_edit.Value == "" or not IsNumber(copies_edit.Value)
            copies_edit.Value := cqtd
        if copies_edit.Value < 1
            copies_edit.Value := 1
        copies_edit.Value := Floor(copies_edit.Value)

        loop copies_edit.Value{
            obj.ActiveSheet.PrintOut(1, 1, 1, ,config_ini["config", "printer"])
        }
        obj.ActiveWorkbook.Saved := true
        obj.WorkBooks.Close()
        obj.Quit()
        Sleep(1000)
        SetTimer(PrintNewCoupons, config_ini["config", "timer_interval", 1000])
        PrintGui.Destroy()
        return
    }
    
    PrintGui    .AddText("yp", "Copias: ")
    copies_edit := PrintGui.AddEdit("yp w30", cqtd)
    
    PrintGui    .Show()
}

Class Coupon{
    __New() {
        this.file_name := ""
        this.code := ""
        this.name := ""
        this.tel := ""
        this.end := ""
        this.dist := ""
        this.city := ""
        this.value := ""
        this.payment_type := ""
        this.is_sale := false
    }

    _GetFromSale(saleTxt){
        this.file_name := StrSplit(saleTxt, "\")
        this.file_name := this.file_name[this.file_name.Length]
        this.code := StrSplit(this.file_name, ".")
        this.code := StrSplit(this.code[1], "_")
        this.code := this.code[2]

        loop read saleTxt{
            local line := A_LoopReadLine

            if InStr(line, "Cliente") and InStr(line, ":") and InStr(line, "/"){
                local cli := StrSplit(line, "/")
                this.name := Trim(cli[2])
                continue
            }
            if InStr(line, "Endereço") and InStr(line, ":"){
                local end := StrSplit(line, ":")
                this.end := Trim(end[2])
                continue
            }
            if InStr(line, "Bairro") and InStr(line, ":"){
                local dist := StrSplit(line, ":")
                this.dist := Trim(dist[2])
                continue
            }
            if InStr(line, "Cidade") and InStr(line, ":"){
                local city := StrSplit(line, ":")
                this.city := Trim(city[2])
                continue
            }
            if InStr(line, "Telefone") and InStr(line, ":"){
                local tel := StrSplit(line, ":")
                this.tel := Trim(tel[2])
                continue
            }
            if InStr(line, "Tipo pagto") and InStr(line, ":"){
                local tp := StrSplit(line, ":")
                this.payment_type := Trim(tp[2])
                continue
            }
            if InStr(line, "Total") and InStr(line, ":") and InStr(line, "R$"){
                local value := StrSplit(line, "R$")
                this.value := Trim(value[2])
                this.value := StrReplace(this.value, ",", ".")
                break
            }
        }
        this.is_sale := true
    }

    _GetFromReceive(receiveTxt){
        file_name := StrSplit(receiveTxt, "\")
        this.file_name := file_name[file_name.Length]
        this.code := StrSplit(this.file_name, ".")
        this.code := StrSplit(this.code[1], "_")
        this.code := this.code[2]

        loop read receiveTxt{
            local line := A_LoopReadLine

            if InStr(line, "Nome do cliente"){
                local cli := StrSplit(line, ".....")
                this.name := Trim(cli[2])
                continue
            }
            if InStr(line, "Valor Recebido"){
                local value := StrSplit(line, "..........")
                this.value := Trim(value[2])
                this.value := StrReplace(this.value, ",", ".")
                continue
            }
        }
        this.is_sale := false
    }

    GetFromLine(line){
        loop parse line, "CSV"{
            if A_Index == 1{
                this.file_name := A_LoopField
                continue
            }
            if A_Index == 2{
                this.code := A_LoopField
                continue
            }
            if A_Index == 3{
                this.name := A_LoopField
                continue
            }
            if A_Index == 4{
                this.tel := A_LoopField
                continue
            }
            if A_Index == 5{
                this.end := A_LoopField
                continue
            }
            if A_Index == 6{
                this.dist := A_LoopField
                continue
            }
            if A_Index == 7{
                this.city := A_LoopField
                continue
            }
            if A_Index == 8{
                this.value := A_LoopField
                continue
            }
            if A_Index == 9{
                this.payment_type := A_LoopField
                continue
            }
            if A_Index == 10{
                this.is_sale := A_LoopField
                continue
            }
            break
        }
    }

    GetFromFile(file){
        if InStr(file, "Recebimento")
            this._GetFromReceive(file)
        else if InStr(file, "ImpressaoVenda")
            this._GetFromSale(file)
    }

    SaveToLine(save_file){
        m := '"{1}","{2}","{3}","{4}","{5}","{6}","{7}","{8}","{9}","{10}"`n'
        m := Format(m, this.file_name, this.code, this.name, this.tel, this.end, this.dist, this.city, this.value, 
                        this.payment_type, this.is_sale)
        FileAppend(m, save_file, "UTF-8")
    }
}

GetNewCoupon(file_pattern, already_checked_files_path){
    global first_run
    static last_check
    if first_run{
        loop files file_pattern{
            local exist := false

            local file_name := A_LoopFileName
            if FileExist(already_checked_files_path){
                loop read already_checked_files_path{
                    local c := Coupon()
                    c.GetFromLine(A_LoopReadLine)
                    
                    if c.file_name == file_name{
                        exist := true
                        break
                    }
                }
            }
            if exist
                continue
            local c := Coupon()
            c.GetFromFile(A_LoopFileFullPath)
            c.SaveToLine(already_checked_files_path)
        }
        first_run := false
        last_check := A_Now
        return ""
    }
    loop files file_pattern{
        if not IsSet(last_check)
            last_check := 0
        ;if A_LoopFileTimeModified < last_check
        ;   continue

        local exist := false

        local file_name := A_LoopFileName
        if FileExist(already_checked_files_path){
            loop read already_checked_files_path{
                local c := Coupon()
                c.GetFromLine(A_LoopReadLine)
                
                if c.file_name == file_name{
                    exist := true
                    break
                }
            }
        }
        if exist
            continue
        c := Coupon()
        c.GetFromFile(A_LoopFileFullPath)
        c.SaveToLine(already_checked_files_path)
        last_check := A_Now
        return c
    }
}
