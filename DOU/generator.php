<?
include_once($_SERVER["DOCUMENT_ROOT"]."/bitrix/modules/main/include.php");

require_once 'vendor/autoload.php';

if (isset($_POST)) {
    $Request = str_replace(array("\r\n", "\r", "\n"), " ", file_get_contents('php://input'));
    $RequestPar = json_decode($Request, true);
    $block=$RequestPar['block'];
    $element=$RequestPar['element'];
    $history = CIBlockElement::GetProperty($block, $element, "sort", "asc", array("CODE" => "LIST_SOGLASOVANIYA"))->Fetch()['VALUE'];
    $date = CIBlockElement::GetProperty($block, $element, "sort", "asc", array("CODE" => "DATA_PODPISANIYA"))->Fetch()['VALUE'];
    $num = CIBlockElement::GetProperty($block, $element, "sort", "asc", array("CODE" => "NOMER_PRIKAZA"))->Fetch()['VALUE'];
    $preamble = CIBlockElement::GetProperty($block, $element, "sort", "asc", array("CODE" => "PREAMBULA"))->Fetch()['VALUE'];
    $name=CIBlockElement::GetByID($element)->Fetch()['NAME'];
    $CREATED_BY = CUser::GetByID(CIBlockElement::GetByID($element)->Fetch()[CREATED_BY])->Fetch();
    $initiator=$CREATED_BY[LAST_NAME].' '.mb_strcut($CREATED_BY[NAME],1,2).'.'.mb_strcut($CREATED_BY[SECOND_NAME],1,2).'. '.$CREATED_BY[PERSONAL_MOBILE];
    $text=CIBlockElement::GetProperty($block, $element, "sort", "asc", array("CODE" => "TEKST_DOKUMENTA"))->Fetch()['VALUE'];
    $text = explode("\n", $text);
    foreach ($text as $value) {
        $arr_text[][text]=$value;
       }
    //$text = str_replace(array("\n"), "<w:br />", $text);
    $history = explode("\n", $history);
    array_splice($history, count($input)-1);
    array_splice($history, 0, 1);
    $k=0;
    $j=0;
    $n=-1;
    foreach ($history as $str) {
        if ((strpos($str, 'Этап согласования №') !== false)){
        $j++;
        $n=-1;
    $k=0;
        }else{
            $el= explode(" ", $str);
            if (is_numeric(strtotime($el[0]))){
                $k++;
                $n++;
                $rsUsers = (CUser::GetList(($by="ID"), ($order="desc"), Array("NAME" => $el[3],"LAST_NAME" => $el[2],"SECOND_NAME" => $el[4])))->Fetch();
                $array[$j][$n]["sogl#$k"]=$el[2].' '.mb_strcut($el[3],1,2).'.'.mb_strcut($el[4],1,2).'. '.$rsUsers[WORK_POSITION];
                $array[$j][$n]["date_sogl#$k"]="$el[0]";
                $array[$j][$n]["result_sogl#$k"]="$el[6]";
            }else{
                for ($i = 0;($i<count($el)); $i++) {
                    $array[$j][$n]["com_sogl#$k"].=$el[$i].' ';
                }
            }
        }
    };

    $document= new \PhpOffice\PhpWord\TemplateProcessor('./template.docx');

    $document->cloneBlock('block', count($array), true, true);

    for ($i = 0; $i < count($array); $i++) {
        $j=$i+1;
        $values = $array[$i];
        $document->setValue("n#$j",$j);
        $document->cloneRowAndSetValues("sogl#$j", $values);
        $values=array();
    }

    $document->setValue('num',$num);
    $document->setValue('name',$name);
    $document->setValue('preamble',$preamble);
    $document->setValue('date',$date);
    $document->setValue('initiator',$initiator);
    $document->cloneRowAndSetValues("text", $arr_text);
    $document->setValue('history',$history);
    $document->saveAs('Приказ.docx'); 

    $way=$_SERVER["DOCUMENT_ROOT"]."/api/DOU/Приказ.docx";
    $file = CFile::MakeFileArray($way);
    CIBlockElement::SetPropertyValueCode($element,'PRIKAZ', $file);
}
    ?>

