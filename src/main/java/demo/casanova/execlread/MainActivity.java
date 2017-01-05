package demo.casanova.execlread;

import android.content.Context;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.util.Xml;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.BaseAdapter;
import android.widget.ListView;
import android.widget.SimpleAdapter;
import android.widget.TextView;

import org.json.JSONArray;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;


public class MainActivity extends AppCompatActivity {

    private ListView infoListView;
    private List<Map<String,String>> list = null;
    private MyAdapter adapter;
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        findView();
        list=new ArrayList();
        adapter =new MyAdapter(this,list,R.layout.listview_item);
        infoListView.setAdapter(adapter);
        // 读取execl文档内容
        String path = Environment.getExternalStorageDirectory()+File.separator+"DCIM";

        XLSXParse.Builder builder =new XLSXParse.Builder()
                .setArmFilePath(path + File.separator + "test.xlsx")
                .setOutFileType(XLSXParse.OutFileType.FILE_TYPE_LIST);
        XLSXParse xl=builder.build();
        Object data= (Object) xl.parseFile();
        adapter.refreshData((List<Map<String,String>>)data);
        JSONArray jsonArray=new JSONArray((List)data);
        Log.i("Main","----"+jsonArray.toString());
    }

    private void findView() {

        infoListView = (ListView) findViewById(R.id.act_listView);
    }

    public class MyAdapter extends BaseAdapter{

        private Context _context;
        private List<Map<String,String>> _list;
        private int _layoutId;

        public MyAdapter(Context context,List<Map<String,String>> list,int layoutId){

            this._context=context;
            this._list=list;
            this._layoutId=layoutId;
        }

        public void refreshData(List<Map<String,String>> list){

            this._list=list;
            this.notifyDataSetChanged();
        }

        @Override
        public int getCount() {
            return this._list.size();
        }

        @Override
        public Object getItem(int position) {
            return this._list.get(position);
        }

        @Override
        public long getItemId(int position) {
            return position;
        }

        @Override
        public View getView(int position, View convertView, ViewGroup parent) {

            ViewHolder holder=null;
            if(convertView==null){
                holder=new ViewHolder();
                convertView = LayoutInflater.from(this._context).inflate(this._layoutId,null);
                holder.txt =(TextView)convertView.findViewById(R.id.act_listView_content);
                convertView.setTag(holder);
            }else{
                holder =(ViewHolder) convertView.getTag();
            }
            holder.txt.setText(getItem(position).toString());
            return convertView;
        }


        private class ViewHolder{
            TextView txt;
        }
    }

}
