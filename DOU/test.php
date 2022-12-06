<?
include_once($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include.php");

require_once 'vendor/autoload.php';

if (isset($_POST)) {
    CIBlockElement::SetPropertyValueCode(3316,'TEST','НОРМ');
    $Request = str_replace(array("\r\n", "\r", "\n"), " ", file_get_contents('php://input'));
    $RequestPar = json_decode($Request, true);
    $block=$RequestPar['block'];
    $element=$RequestPar['element'];
    $history = CIBlockElement::GetProperty($block, $element, "sort", "asc", array("CODE" => "LIST_SOGLASOVANIYA"))->Fetch()['VALUE'];
    $date = CIBlockElement::GetProperty($block, $element, "sort", "asc", array("CODE" => "DATA_PODPISANIYA"))->Fetch()['VALUE'];
    $num = CIBlockElement::GetProperty($block, $element, "sort", "asc", array("CODE" => "NOMER_PRIKAZA"))->Fetch()['VALUE'];
    $preamble = CIBlockElement::GetProperty($block, $element, "sort", "asc", array("CODE" => "PREAMBULA"))->Fetch()['VALUE'];
    $name=CIBlockElement::GetByID($element)->Fetch()['NAME'];
    //$text=CIBlockElement::GetProperty($block, $element, "sort", "asc", array("CODE" => "TEKST"))->Fetch()['VALUE']['TEXT'];
    $text=CIBlockElement::GetProperty($block, $element, "sort", "asc", array("CODE" => "TEKST_DOKUMENTA"))->Fetch()['VALUE'];
    $text = str_replace(array("\n"), "<w:br />", $text);
    $history = explode("\n", $history);
    //array_splice($history, count($history)-1, 1);

    $replacements=array();
    $j=1;
    $k=0;
    foreach ($history as $str) {
        if (!(strpos($str, 'Этап согласования №') !== false)){
            $el= explode(" ", $str);
        if ($el[0]=="Комментарий:"){
            for ($i = 1;($i<count($el)); $i++) {
                $com_sogl.=$el[$i].' ';
             }
            $arr["com_sogl$j"]="$com_sogl";
            $com_sogl='';
            $j++;
        }else{
            $rsUsers = (CUser::GetList(($by="ID"), ($order="desc"), Array("NAME" => $el[3],"LAST_NAME" => $el[2],"SECOND_NAME" => $el[4])))->Fetch();
            $arr["sogl$j"]=$el[2].' '.mb_strcut($el[3],1,2).'.'.mb_strcut($el[4],1,2).'.<w:br />'.$rsUsers[WORK_POSITION];
            $arr["date_sogl$j"]="$el[0]";
            $arr["result_sogl$j"]="$el[6]";
            }
        }else{	
            $j=1;
            $replacements[]=$arr;
            $k++;
            $arr = array("n"=>"$k","sogl1" => "-", "date_sogl1" => "-", "result_sogl1" => "-", "com_sogl1" => "-","sogl2" => "-", "date_sogl2" => "-", "result_sogl2" => "-", "com_sogl2" => "-","sogl3" => "-", "date_sogl3" => "-", "result_sogl3" => "-", "com_sogl3" => "-");
        }
    }
    $replacements[]=$arr;
    array_splice($replacements, 0, 1);

    $document= new \PhpOffice\PhpWord\TemplateProcessor('./template.docx');
    $document->cloneBlock('block_name', 0, true, false, $replacements);
    $document->setValue('num',$num);
    $document->setValue('name',$name);
    $document->setValue('preamble',$preamble);
    $document->setValue('date',$date);
    $document->setValue('text',$text);
    $document->setValue('history',$history);
    $document->saveAs('full_template.docx'); 

    $way=$_SERVER["DOCUMENT_ROOT"]."/api/DOU/full_template.docx";
    $file = CFile::MakeFileArray($way);
    CIBlockElement::SetPropertyValueCode($element,'PRIKAZ', $file);
}
    ?>

