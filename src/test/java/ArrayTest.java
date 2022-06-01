import java.util.ArrayList;
import java.util.List;

public class ArrayTest {

    public static void main(String[] args) {
        ArrayList<String> list = new ArrayList<>();
        list.add("1");
        list.add("2");
        list.add("3");

        List<String> subList = list.subList(1, 3);
        System.out.println(subList);
        for (int i = 0; i < list.size(); i++) {
            System.out.println(list.get(i));
            i++;
        }
    }

}
