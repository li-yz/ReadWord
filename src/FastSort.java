import java.util.ArrayList;
import java.util.List;

/**
 * Created by Liyanzhen on 2016/12/28.
 */
public class FastSort {

    public static void main(String[] args){
        List<Integer> list = new ArrayList<Integer>();
        list.add(5);
        list.add(6);
        list.add(3);
        list.add(7);
        list.add(4);
        list.add(2);

        fastSort(list,0,list.size()-1);
        MyPrint.print(""+list);
    }

    public static void fastSort(List<Integer> list , int start , int end){
        if(start < end){
            int index = partition(list,start,end);

            fastSort(list,start,index-1);
            fastSort(list,index+1,end);
        }
    }

    private static int partition(List<Integer> list, int low, int high){
        int temp=list.get(low);
        while(low < high){//当low==high时 退出循环
            while(low < high && list.get(high) >= temp ){
                high--;
            }
            list.set(low,list.get(high));
            while(low < high && list.get(low) <= temp){
                low++;
            }
            list.set(high,list.get(low));

        }
        list.set(low,temp);
        return low;
    }
}
