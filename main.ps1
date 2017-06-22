###
# UtilsFunction
###
function isEmpty($str){
  if($str.length -le 0){
    return $true
  }
  else {
    return $false
  }
}

function exist($dest){
  if(Test-Path -path ([System.io.Path]::GetFullPath("${dest}"))){
    return $true
  }
  else{
    return $false
  }
}

function worksheetExists($b, $name){
  $hasWorkSheet = $false
  foreach($s in $b.WorkSheets){
    echo $s.Name
    if($s.Name -eq $name){
      $hasWorkSheet = $true
      break
    }
  }
  return $hasWorkSheet
}

function px2cm($px, $margin=10){
  $inch = 2.54
  $dpi  = 220.0
  return [Double](($px / $dpi * $inch) * 10) + $margin
}

function getRGB($r, $g, $b){
  return ($r + $g * 256 + $b * 256 * 256)
}


###
# var
###

# book info
$bookName   = "master.xlsm"
$bookPath   = "${HOME}\Desktop\${bookName}"
$importPath = ([System.io.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)) + "\src"

# xlConst
$xlDouble                      = -4119
$xlDot                         = -4118
$xlDash                        = -4115
$xlLineStyleNone               = -4142
$xlContinuous                  = 1
$xlDashDot                     = 4
$xlDashDotDot                  = 5
$xlSlantDashDot                = 13
$xlValidateList                = 3
$xlValidAlertStop              = 1
$xlEqual                       = 3
$xlCellValue                   = 1
$xlBetween                     = 1
$xlOpenXMLWorkbookMacroEnabled = 52
$xlHairline                    = 1

# color
$white     = getRGB 255 255 255
$darkGrey  = getRGB 64 64 64
$lightGrey = getRGB 166 166 166
$green     = getRGB 0 204 153

# sheetName
$memberSheetName  = "member"
$statusSheetName  = "status"
$holidaySheetName = "holiday"
$mainSheetName    = "main"

# dataSheet data
$memberHead  = @("id", "value", "color", "description")
$memberColor = @(
  "192,226,230", "253,207,148", "223,220,213",
  "242,240,218", "251,202,210", "167,237,169",
  "228,211,241", "255,172,183", "254,242,139",
  "154,231,217"
)
$memberVals = @(
  "member0", "member1", "member2", "member3", "member4",
  "member5", "member6", "member7", "member8", "member9"
)

$statusHead = @("id", "value", "color", "description")
$statusVals = @(
  "未着手", "準備中", "対応中",
  "制作中", "確認中", "テスト中",
  "承認中", "対応済", "保留",
  "リリース", "中止"
)
$statusColor = @(
  "255,255,255", "242,242,242", "230,230,230",
  "204,204,204", "179,179,179", "153,153,153",
  "128,128,128", "102,102,102", "77,77,77",
  "51,51,51", "26,26,26"
)

$holidayHead = @("id", "day", "day week", "description")
$holidayVals = @(
  "2013/1/1,火,元日",
  "2013/1/14,月,成人の日",
  "2013/2/11,月,建国記念の日",
  "2013/3/20,水,春分の日",
  "2013/4/29,月,昭和の日",
  "2013/5/3,金,憲法記念日",
  "2013/5/4,土,みどりの日",
  "2013/5/5,日,こどもの日",
  "2013/5/6,月,振替休日",
  "2013/7/15,月,海の日",
  "2013/9/16,月,敬老の日",
  "2013/9/23,月,秋分の日",
  "2013/10/14,月,体育の日",
  "2013/11/3,日,文化の日",
  "2013/11/4,月,振替休日",
  "2013/11/23,土,勤労感謝の日",
  "2013/12/23,月,天皇誕生日",
  "2014/1/1,水,元日",
  "2014/1/13,月,成人の日",
  "2014/2/11,火,建国記念の日",
  "2014/3/21,金,春分の日",
  "2014/4/29,火,昭和の日",
  "2014/5/3,土,憲法記念日",
  "2014/5/4,日,みどりの日",
  "2014/5/5,月,こどもの日",
  "2014/5/6,火,振替休日",
  "2014/7/21,月,海の日",
  "2014/9/15,月,敬老の日",
  "2014/9/23,火,秋分の日",
  "2014/10/13,月,体育の日",
  "2014/11/3,月,文化の日",
  "2014/11/23,日,勤労感謝の日",
  "2014/11/24,月,振替休日",
  "2014/12/23,火,天皇誕生日",
  "2015/1/1,木,元日",
  "2015/1/12,月,成人の日",
  "2015/2/11,水,建国記念の日",
  "2015/3/21,土,春分の日",
  "2015/4/29,水,昭和の日",
  "2015/5/3,日,憲法記念日",
  "2015/5/4,月,みどりの日",
  "2015/5/5,火,こどもの日",
  "2015/5/6,水,振替休日",
  "2015/7/20,月,海の日",
  "2015/9/21,月,敬老の日",
  "2015/9/22,火,国民の休日",
  "2015/9/23,水,秋分の日",
  "2015/10/12,月,体育の日",
  "2015/11/3,火,文化の日",
  "2015/11/23,月,勤労感謝の日",
  "2015/12/23,水,天皇誕生日",
  "2016/1/1,金,元日",
  "2016/1/11,月,成人の日",
  "2016/2/11,木,建国記念の日",
  "2016/3/20,日,春分の日",
  "2016/3/21,月,振替休日",
  "2016/4/29,金,昭和の日",
  "2016/5/3,火,憲法記念日",
  "2016/5/4,水,みどりの日",
  "2016/5/5,木,こどもの日",
  "2016/7/18,月,海の日",
  "2016/8/11,木,山の日",
  "2016/9/19,月,敬老の日",
  "2016/9/22,木,秋分の日",
  "2016/10/10,月,体育の日",
  "2016/11/3,木,文化の日",
  "2016/11/23,水,勤労感謝の日",
  "2016/12/23,金,天皇誕生日",
  "2017/1/1,日,元日",
  "2017/1/2,月,振替休日",
  "2017/1/9,月,成人の日",
  "2017/2/11,土,建国記念の日",
  "2017/3/20,月,春分の日",
  "2017/4/29,土,昭和の日",
  "2017/5/3,水,憲法記念日",
  "2017/5/4,木,みどりの日",
  "2017/5/5,金,こどもの日",
  "2017/7/17,月,海の日",
  "2017/8/11,金,山の日",
  "2017/9/18,月,敬老の日",
  "2017/9/23,土,秋分の日",
  "2017/10/9,月,体育の日",
  "2017/11/3,金,文化の日",
  "2017/11/23,木,勤労感謝の日",
  "2017/12/23,土,天皇誕生日",
  "2018/1/1,月,元日",
  "2018/1/8,月,成人の日",
  "2018/2/11,日,建国記念の日",
  "2018/2/12,月,振替休日",
  "2018/3/21,水,春分の日",
  "2018/4/29,日,昭和の日",
  "2018/4/30,月,振替休日",
  "2018/5/3,木,憲法記念日",
  "2018/5/4,金,みどりの日",
  "2018/5/5,土,こどもの日",
  "2018/7/16,月,海の日",
  "2018/8/11,土,山の日",
  "2018/9/17,月,敬老の日",
  "2018/9/23,日,秋分の日",
  "2018/9/24,月,振替休日",
  "2018/10/8,月,体育の日",
  "2018/11/3,土,文化の日",
  "2018/11/23,金,勤労感謝の日",
  "2018/12/23,日,天皇誕生日",
  "2018/12/24,月,振替休日",
  "2019/1/1,火,元日",
  "2019/1/14,月,成人の日",
  "2019/2/11,月,建国記念の日",
  "2019/3/21,木,春分の日",
  "2019/4/29,月,昭和の日",
  "2019/5/3,金,憲法記念日",
  "2019/5/4,土,みどりの日",
  "2019/5/5,日,こどもの日",
  "2019/5/6,月,振替休日",
  "2019/7/15,月,海の日",
  "2019/8/11,日,山の日",
  "2019/8/12,月,振替休日",
  "2019/9/16,月,敬老の日",
  "2019/9/23,月,秋分の日",
  "2019/10/14,月,体育の日",
  "2019/11/3,日,文化の日",
  "2019/11/4,月,振替休日",
  "2019/11/23,土,勤労感謝の日",
  "2019/12/23,月,天皇誕生日",
  "2020/1/1,水,元日",
  "2020/1/13,月,成人の日",
  "2020/2/11,火,建国記念の日",
  "2020/3/20,金,春分の日",
  "2020/4/29,水,昭和の日",
  "2020/5/3,日,憲法記念日",
  "2020/5/4,月,みどりの日",
  "2020/5/5,火,こどもの日",
  "2020/5/6,水,振替休日",
  "2020/7/20,月,海の日",
  "2020/8/11,火,山の日",
  "2020/9/21,月,敬老の日",
  "2020/9/22,火,秋分の日",
  "2020/10/12,月,体育の日",
  "2020/11/3,火,文化の日",
  "2020/11/23,月,勤労感謝の日",
  "2020/12/23,水,天皇誕生日",
  "2021/1/1,金,元日",
  "2021/1/11,月,成人の日",
  "2021/2/11,木,建国記念の日",
  "2021/3/20,土,春分の日",
  "2021/4/29,木,昭和の日",
  "2021/5/3,月,憲法記念日",
  "2021/5/4,火,みどりの日",
  "2021/5/5,水,こどもの日",
  "2021/7/19,月,海の日",
  "2021/8/11,水,山の日",
  "2021/9/20,月,敬老の日",
  "2021/9/23,木,秋分の日",
  "2021/10/11,月,体育の日",
  "2021/11/3,水,文化の日",
  "2021/11/23,火,勤労感謝の日",
  "2021/12/23,木,天皇誕生日",
  "2022/1/1,土,元日",
  "2022/1/10,月,成人の日",
  "2022/2/11,金,建国記念の日",
  "2022/3/21,月,春分の日",
  "2022/4/29,金,昭和の日",
  "2022/5/3,火,憲法記念日",
  "2022/5/4,水,みどりの日",
  "2022/5/5,木,こどもの日",
  "2022/7/18,月,海の日",
  "2022/8/11,木,山の日",
  "2022/9/19,月,敬老の日",
  "2022/9/23,金,秋分の日",
  "2022/10/10,月,体育の日",
  "2022/11/3,木,文化の日",
  "2022/11/23,水,勤労感謝の日",
  "2022/12/23,金,天皇誕生日",
  "2023/1/1,日,元日",
  "2023/1/2,月,振替休日",
  "2023/1/9,月,成人の日",
  "2023/2/11,土,建国記念の日",
  "2023/3/21,火,春分の日",
  "2023/4/29,土,昭和の日",
  "2023/5/3,水,憲法記念日",
  "2023/5/4,木,みどりの日",
  "2023/5/5,金,こどもの日",
  "2023/7/17,月,海の日",
  "2023/8/11,金,山の日",
  "2023/9/18,月,敬老の日",
  "2023/9/23,土,秋分の日",
  "2023/10/9,月,体育の日",
  "2023/11/3,金,文化の日",
  "2023/11/23,木,勤労感謝の日",
  "2023/12/23,土,天皇誕生日",
  "2024/1/1,月,元日",
  "2024/1/8,月,成人の日",
  "2024/2/11,日,建国記念の日",
  "2024/2/12,月,振替休日",
  "2024/3/20,水,春分の日",
  "2024/4/29,月,昭和の日",
  "2024/5/3,金,憲法記念日",
  "2024/5/4,土,みどりの日",
  "2024/5/5,日,こどもの日",
  "2024/5/6,月,振替休日",
  "2024/7/15,月,海の日",
  "2024/8/11,日,山の日",
  "2024/8/12,月,振替休日",
  "2024/9/16,月,敬老の日",
  "2024/9/22,日,秋分の日",
  "2024/9/23,月,振替休日",
  "2024/10/14,月,体育の日",
  "2024/11/3,日,文化の日",
  "2024/11/4,月,振替休日",
  "2024/11/23,土,勤労感謝の日",
  "2024/12/23,月,天皇誕生日",
  "2025/1/1,水,元日",
  "2025/1/13,月,成人の日",
  "2025/2/11,火,建国記念の日",
  "2025/3/20,木,春分の日",
  "2025/4/29,火,昭和の日",
  "2025/5/3,土,憲法記念日",
  "2025/5/4,日,みどりの日",
  "2025/5/5,月,こどもの日",
  "2025/5/6,火,振替休日",
  "2025/7/21,月,海の日",
  "2025/8/11,月,山の日",
  "2025/9/15,月,敬老の日",
  "2025/9/23,火,秋分の日",
  "2025/10/13,月,体育の日",
  "2025/11/3,月,文化の日",
  "2025/11/23,日,勤労感謝の日",
  "2025/11/24,月,振替休日",
  "2025/12/23,火,天皇誕生日",
  "2026/1/1,木,元日",
  "2026/1/12,月,成人の日",
  "2026/2/11,水,建国記念の日",
  "2026/3/20,金,春分の日",
  "2026/4/29,水,昭和の日",
  "2026/5/3,日,憲法記念日",
  "2026/5/4,月,みどりの日",
  "2026/5/5,火,こどもの日",
  "2026/5/6,水,振替休日",
  "2026/7/20,月,海の日",
  "2026/8/11,火,山の日",
  "2026/9/21,月,敬老の日",
  "2026/9/22,火,国民の休日",
  "2026/9/23,水,秋分の日",
  "2026/10/12,月,体育の日",
  "2026/11/3,火,文化の日",
  "2026/11/23,月,勤労感謝の日",
  "2026/12/23,水,天皇誕生日",
  "2027/1/1,金,元日",
  "2027/1/11,月,成人の日",
  "2027/2/11,木,建国記念の日",
  "2027/3/21,日,春分の日",
  "2027/3/22,月,振替休日",
  "2027/4/29,木,昭和の日",
  "2027/5/3,月,憲法記念日",
  "2027/5/4,火,みどりの日",
  "2027/5/5,水,こどもの日",
  "2027/7/19,月,海の日",
  "2027/8/11,水,山の日",
  "2027/9/20,月,敬老の日",
  "2027/9/23,木,秋分の日",
  "2027/10/11,月,体育の日",
  "2027/11/3,水,文化の日",
  "2027/11/23,火,勤労感謝の日",
  "2027/12/23,木,天皇誕生日",
  "2028/1/1,土,元日",
  "2028/1/10,月,成人の日",
  "2028/2/11,金,建国記念の日",
  "2028/3/20,月,春分の日",
  "2028/4/29,土,昭和の日",
  "2028/5/3,水,憲法記念日",
  "2028/5/4,木,みどりの日",
  "2028/5/5,金,こどもの日",
  "2028/7/17,月,海の日",
  "2028/8/11,金,山の日",
  "2028/9/18,月,敬老の日",
  "2028/9/22,金,秋分の日",
  "2028/10/9,月,体育の日",
  "2028/11/3,金,文化の日",
  "2028/11/23,木,勤労感謝の日",
  "2028/12/23,土,天皇誕生日",
  "2029/1/1,月,元日",
  "2029/1/8,月,成人の日",
  "2029/2/11,日,建国記念の日",
  "2029/2/12,月,振替休日",
  "2029/3/20,火,春分の日",
  "2029/4/29,日,昭和の日",
  "2029/4/30,月,振替休日",
  "2029/5/3,木,憲法記念日",
  "2029/5/4,金,みどりの日",
  "2029/5/5,土,こどもの日",
  "2029/7/16,月,海の日",
  "2029/8/11,土,山の日",
  "2029/9/17,月,敬老の日",
  "2029/9/23,日,秋分の日",
  "2029/9/24,月,振替休日",
  "2029/10/8,月,体育の日",
  "2029/11/3,土,文化の日",
  "2029/11/23,金,勤労感謝の日",
  "2029/12/23,日,天皇誕生日",
  "2029/12/24,月,振替休日",
  "2030/1/1,火,元日",
  "2030/1/14,月,成人の日",
  "2030/2/11,月,建国記念の日",
  "2030/3/20,水,春分の日",
  "2030/4/29,月,昭和の日",
  "2030/5/3,金,憲法記念日",
  "2030/5/4,土,みどりの日",
  "2030/5/5,日,こどもの日",
  "2030/5/6,月,振替休日",
  "2030/7/15,月,海の日",
  "2030/8/11,日,山の日",
  "2030/8/12,月,振替休日",
  "2030/9/16,月,敬老の日",
  "2030/9/23,月,秋分の日",
  "2030/10/14,月,体育の日",
  "2030/11/3,日,文化の日",
  "2030/11/4,月,振替休日",
  "2030/11/23,土,勤労感謝の日",
  "2030/12/23,月,天皇誕生日",
  "2031/1/1,水,元日",
  "2031/1/13,月,成人の日",
  "2031/2/11,火,建国記念の日",
  "2031/3/21,金,春分の日",
  "2031/4/29,火,昭和の日",
  "2031/5/3,土,憲法記念日",
  "2031/5/4,日,みどりの日",
  "2031/5/5,月,こどもの日",
  "2031/5/6,火,振替休日",
  "2031/7/21,月,海の日",
  "2031/8/11,月,山の日",
  "2031/9/15,月,敬老の日",
  "2031/9/23,火,秋分の日",
  "2031/10/13,月,体育の日",
  "2031/11/3,月,文化の日",
  "2031/11/23,日,勤労感謝の日",
  "2031/11/24,月,振替休日",
  "2031/12/23,火,天皇誕生日",
  "2032/1/1,木,元日",
  "2032/1/12,月,成人の日",
  "2032/2/11,水,建国記念の日",
  "2032/3/20,土,春分の日",
  "2032/4/29,木,昭和の日",
  "2032/5/3,月,憲法記念日",
  "2032/5/4,火,みどりの日",
  "2032/5/5,水,こどもの日",
  "2032/7/19,月,海の日",
  "2032/8/11,水,山の日",
  "2032/9/20,月,敬老の日",
  "2032/9/21,火,国民の休日",
  "2032/9/22,水,秋分の日",
  "2032/10/11,月,体育の日",
  "2032/11/3,水,文化の日",
  "2032/11/23,火,勤労感謝の日",
  "2032/12/23,木,天皇誕生日"
)


###
# workbookFunction
###
function namesInit($b){
  $b.Names.Add("RTVStartDate", (Get-Date).ToString("yyyy/MM/dd")) | Out-Null
  $b.Names.Add("RTVEndDate", (Get-Date).AddMonths(2).ToString("yyyy/MM/dd")) | Out-Null
}

function dataSheetInit($b, $sheetName, $head, $color, $vals){
  $rowSize      = $vals.length
  $hasWorkSheet = worksheetExists $sheetName $b

  if($hasWorkSheet -eq $false){
    $sheet = $b.WorkSheets.Add()
    $sheet.Name      = $sheetName
    $sheet.Tab.Color = $green

    # FreezePanes
    $sheet.Active
    $sheet.Range("A2").Select() | Out-Null
    $excel.ActiveWindow.FreezePanes = $true

    # head
    $h                = $sheet.Range("A1", $sheet.Cells.Item(1, $sheet.Columns.Count))
    $h.Interior.Color = $darkGrey
    $h.Font.Color     = $white
    $h.Font.Bold      = $true

    for($i = 0; $i -lt $head.length; $i++){
      $s                   = $sheet.Cells.Item(1, $i + 1)
      $s.Value             = $head[$i]
      $s.Borders.LineStyle = $xlDash
      $s.Borders.Weight    = $xlHairline
    }

    for($i = 0; $i -lt $rowSize; $i++){
      $row = $i + 2
      $sheet.Cells.Item($row, 1).Value = [String]($i + 1)

      if($sheetName -ne $holidaySheetName){
        $sheet.Cells.Item($row, 2).Value = $vals[$i]
        $sheet.Cells.Item($row, 3).Value = $color[$i]

        $r = [Int](($color[$i] -split ',')[0])
        $g = [Int](($color[$i] -split ',')[1])
        $b = [Int](($color[$i] -split ',')[2])
        $sheet.Cells($i + 2, 3).Interior.Color = (getRGB $r $g $b)
      }
      else{
        $v = ($vals[$i] -split ',')
        $sheet.Cells.Item($row, 2).Value              = $v[0]
        $sheet.Cells.Item($row, 2).Offset(0, 1).Value = $v[1]
        $sheet.Cells.Item($row, 2).Offset(0, 2).Value = $v[2]
      }

      if(($sheetName -eq $statusSheetName) -and $i -ge 5){
        $sheet.Cells($i + 2, 3).Font.Color = $white
      }

      foreach($a in 1..4){
        $sheet.Cells.Item($row, $a).Borders.LineStyle = $xlDash
        $sheet.Cells.Item($row, $a).Borders.Weight    = $xlHairline
      }

    }
    return $sheet
  }
}

function mainSheetInit($b, $sheetName){
  $head = @(
    "備考欄", "No.", "タスク", "担当", "開始日",
    "作業日数", "調整日数", "完了日", "ステータス"
  )

  $tasks = @(
    "入稿", "確認", "制作", "テストアップ", "確認",
    "修正", "確認", "修正", "確認、認証", "公開"
  )
  $taskName = "taskName"

  $hasWorkSheet = worksheetExists $sheetName $b

  if($hasWorkSheet -eq $false){
    $sheet      = $b.WorkSheets.Add()
    $sheet.Name = $sheetName

    # FreezePanes
    $sheet.Active
    $sheet.Range("J6").Select() | Out-Null
    $excel.ActiveWindow.FreezePanes = $true

    # Title
    $sheet.Range("A1").Font.Bold = $true
    $sheet.Range("A1").Value     = "ProjectName"

    # memberSample
    $sheet.Range("J1").Value = '=member!$B$2'
    $sheet.Range("J2").Value = '=member!$B$6'
    $sheet.Range("N1").Value = '=member!$B$3'
    $sheet.Range("N2").Value = '=member!$B$8'
    $sheet.Range("R1").Value = '=member!$B$4'
    $sheet.Range("R2").Value = '=member!$B$9'
    $sheet.Range("V1").Value = '=member!$B$5'
    $sheet.Range("V2").Value = '=member!$B$10'
    $sheet.Range("Z1").Value = '=member!$B$6'
    $sheet.Range("Z2").Value = '=member!$B$11'

    # bold
    $sheet.Range($sheet.Cells.Item(6, 1), $sheet.Cells.Item($sheet.Rows.Count, $sheet.Columns.Count)).Borders.LineStyle = $xlDot
    $sheet.Range($sheet.Cells.Item(6, 1), $sheet.Cells.Item($sheet.Rows.Count, $sheet.Columns.Count)).Borders.Weight    = $xlHairline

    # head
    for($i=0; $i -lt $head.length; $i++){
      $sheet.Cells.Item(4, $i + 1).Value     = $head[$i]
      $sheet.Cells.Item(4, $i + 1).Font.Bold = $true
    }

    # dateArea
    # paint bg to darkGrey, color to white
    $sheet.Range($sheet.Rows(3), $sheet.Rows(5)).Interior.Color = $darkGrey
    $sheet.Range($sheet.Rows(3), $sheet.Rows(5)).Font.Color     = $white

    # member formatConditions
    for($i=0; $i -lt $memberVals.length; $i++){
      $sheet.Range('$1:$2').FormatConditions.Add($xlCellValue, $xlEqual, '=INDIRECT(CONCATENATE("member!","$B","$' + ($i + 2) + '"))').Interior.Color = [Long]($book.WorkSheets($memberSheetName).Cells.Item($i + 2, 3).Interior.Color)
      $sheet.Range($sheet.Cells.Item(7, 4), $sheet.Cells.Item($sheet.Rows.Count, 4)).FormatConditions.Add($xlCellValue, $xlEqual, '=INDIRECT(CONCATENATE("member!","$B","$' + ($i + 2) + '"))').Interior.Color = [Long]($book.WorkSheets($memberSheetName).Cells.Item($i + 2, 3).Interior.Color)
    }

    # status formatConditions
    for($i=0; $i -lt $statusVals.length; $i++){
      $f = $sheet.Range($sheet.Cells.Item(7, 9), $sheet.Cells.Item($sheet.Rows.Count, 9)).FormatConditions.Add($xlCellValue, $xlEqual, '=INDIRECT(CONCATENATE("status!","$B","$' + ($i + 2) + '"))')
      if($i -gt 3){
        $f.Interior.Color = [Long]($book.WorkSheets($statusSheetName).Cells.Item($i + 2, 3).Interior.Color)
        $f.Font.Color     = $white
      }
      else{
        $f.Interior.Color = [Long]($book.WorkSheets($statusSheetName).Cells.Item($i + 2, 3).Interior.Color)
      }
    }

    # print task
    # task head
    $taskSize = 3
    $cRow     = 6
    for($i=0; $i -lt $taskSize; $i++){
      # taskHead
      $sheet.Range($sheet.Rows($cRow), $sheet.Rows($cRow)).Interior.Color = $lightGrey
      $sheet.Range($sheet.Rows($cRow), $sheet.Rows($cRow)).Font.Color     = $white
      $sheet.Range($sheet.Rows($cRow), $sheet.Rows($cRow)).Font.Bold      = $true
      $sheet.Cells.Item($cRow, 3).Value                                   = $taskName
      $cRow = $cRow + 1

      for($j=0; $j -lt $tasks.length; $j++){
        if(($i -eq 0) -and ($j -eq 0)){
          # taskId
          $sheet.Cells.Item($cRow, 2).Value = 1
          # startDate
          $sheet.Cells.Item($cRow, 5).Value = (Get-Date).ToString("yyyy/MM/dd")
        }
        else{
          # taskId
          $sheet.Cells.Item($cRow, 2).Value = '=nextId(curtAddr())'
          # startDate
          $sheet.Cells.Item($cRow, 5).Value = '=getStartDate()'
          $sheet.Cells.Item($cRow, 5).NumberFormatLocal = "yyyy/MM/dd"
        }

        # task name
        $sheet.Cells.Item($cRow, 3).Value = $tasks[$j]
        # task name indent
        if(($j -eq 2) -or ($j -ge 4 -and $j -le 9)){
          $sheet.Cells.Item($cRow, 3).IndentLevel = 1
        }
        # member
        $sheet.Cells.Item($cRow, 4).Value          = $book.WorkSheets($memberSheetName).Cells.Item($j + 2, 2).Value2
        $sheet.Cells.Item($cRow, 4).Interior.Color = [Long]($book.WorkSheets($memberSheetName).Cells.Item($j + 2, 3).Interior.Color)
        $sheet.Cells.Item($cRow, 4).Validation.Add($xlValidateList, $xlValidAlertStop, $xlEqual, '=INDIRECT(CONCATENATE("member!","$B$2",":","$B$100"))')
        # workday
        $sheet.Cells.Item($cRow, 6).Value = 1
        # spendday
        $sheet.Cells.Item($cRow, 7).Value = 1
        # endDay
        $sheet.Cells.Item($cRow, 8).Value             = '=getEndDate()'
        $sheet.Cells.Item($cRow, 8).NumberFormatLocal = "yyyy/MM/dd"
        # status
        $sheet.Cells.Item($cRow, 9).Value          = $book.WorkSheets($statusSheetName).Cells.Item($j + 2, 2).Text
        $sheet.Cells.Item($cRow, 9).Interior.Color = [Long]($book.WorkSheets($statusSheetName).Cells.Item($j + 2, 3).Interior.Color)
        $sheet.Cells.Item($cRow, 9).Validation.Add($xlValidateList, $xlValidAlertStop, $xlEqual, '=INDIRECT(CONCATENATE("status!","$B$2",":","$B$100"))')

        $cRow = $cRow + 1
      }
    }
  }
}

function importScript($b){
  $importFiles = @(
    "UtilModule.bas",
    "MainForm.frm", "RVCalendarForm.frm",
    "MainModule.bas", "MainSheet.cls",
    "RVDateItem.cls", "RVCalendar.cls"
  )

  $pjt = $book.VBProject

  # set objectName as same as sheetName + "Sheet".
  foreach($s in $book.WorkSheets){
    $pjt.VBComponents($s.CodeName).Properties("_Codename").Value = $s.Name + "Sheet"
  }

  # import script from src.
  foreach($f in $importFiles){
    $src = "${importPath}\${f}"
    if(exist $src){
      if($f -eq "MainSheet.cls"){
        $pjt.VBComponents("MainSheet").CodeModule.AddFromFile($src)
      }
      else{
        $pjt.VBComponents.Import($src) | Out-Null
      }
    }
    else {
      echo "${src}"
      echo "No such file."
    }
  }
}

###
# main
###
if(exist $bookPath){
  echo "${bookPath} already exists."
  echo "Bye."
  exit 0
}

if(-not(exist $importPath)){
  echo "${importPath}"
  echo "No such Files or Directory."
  echo "Bye."
  exit 0
}

try{
  # new Excel Object
  $excel         = New-Object -ComObject Excel.Application
  $excel.Visible = $false

  # book open
  $book = $excel.Workbooks.Add()

  namesInit $book | Out-Null

  # attach dataSheet
  dataSheetInit $book $statusSheetName $statusHead $statusColor $statusVals | Out-Null
  dataSheetInit $book $memberSheetName $memberHead $memberColor $memberVals | Out-Null
  dataSheetInit $book $holidaySheetName $holidayHead @() $holidayVals | Out-Null

  # remove Sheet1 after build dataSheet.
  $book.WorkSheets("Sheet1").Delete() | Out-Null

  # attach mainSheet
  mainSheetInit $book $mainSheetName | Out-Null

  # import script.
  importScript $book | Out-Null

  # set a shortcut to call the control panel.
  $book.Application.MacroOptions("loadMainForm", "", $false, $false, $true, "e")
}
catch [Exception]{
  echo $error
}
finally{
  # activate screen update
  $excel.Application.ScreenUpdating = $true
  $excel.Application.Calculation    = -4105

  $excel.Application.DisplayAlerts = $false
  $book.SaveAs($bookPath, $xlOpenXMLWorkbookMacroEnabled)
  $excel.Application.DisplayAlerts = $true

  $excel.Quit()
}
