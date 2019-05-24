import alena.ExcelParser_1;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by misch on 11.09.2017.
 */
public class Main {

    public static final Pattern p = Pattern.compile("[а-яА-Я]{4,5}+[\\-]+[0-9]{2}+[\\-]+[0-9]{2}+");

    public static boolean doMatch(String word) {
        Matcher matcher = p.matcher(word);
        return matcher.matches();
    }
    public static void main(String args[]){
        //Чтение зависит от формата: xlsx or xls, поэтому просто кидаю ссылку. Там оба варика чтения есть.
        //http://javadevblog.com/rabotaem-s-excel-v-java.html
        //кратко: создаем выходящий поток через класс файла, читаем файл через POI, нопределяем название листаи пошло-поехало.

        System.out.println(doMatch("БББО-01-16"));
    }
}
