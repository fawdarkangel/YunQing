function varargout = Auto2_4_2(varargin)
% AUTO2_4_2 MATLAB code for Auto2_4_2.fig
%      AUTO2_4_2, by itself, creates a new AUTO2_4_2 or raises the existing
%      singleton*.
%
%      H = AUTO2_4_2 returns the handle to a new AUTO2_4_2 or the handle to
%      the existing singleton*.
%
%      AUTO2_4_2('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in AUTO2_4_2.M with the given input arguments.
%
%      AUTO2_4_2('Property','Value',...) creates a new AUTO2_4_2 or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before Auto2_4_2_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to Auto2_4_2_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help Auto2_4_2

% Last Modified by GUIDE v2.5 25-Mar-2020 12:04:24

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto2_4_2_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto2_4_2_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before Auto2_4_2 is made visible.
function Auto2_4_2_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
movegui(gcf,'center')
b=load([cd,'\interface\Fahrzeugcode.mat']);
for i=1:length(b.Fahrzeugcode)
Fahrzeugcode{i,1}=b.Fahrzeugcode{i,2};
end
set(handles.Fahrzeugcode,'String',Fahrzeugcode);
set(handles.popupmenu3,'Value',2);
set(handles.popupmenu4,'Value',5);
  COLOR_INDEX=[204 0 0;204 189 0;58 204 0;0 204 180;0 49 204;97 0 204;204 0 151;...
      148 71 56;141 148 56;55 149 66;56 148 139;92 72 132;125 73 131;130 74 74;226 130 226;...
      99 23 99;57 231 78;22 188 42]/255;
  setappdata(0,'Auto2_4_2_COLOR_INDEX',COLOR_INDEX);
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto2_4_2 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto2_4_2_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on selection change in listbox1.
function listbox1_Callback(hObject, eventdata, handles)
cla(handles.axes1);
MP=getappdata(0,'Auto2_4_2_MP');
FigureIndex=getappdata(0,'Auto2_4_2_FigureIndex');
MaxKraft=getappdata(0,'Auto2_4_2_MaxKraft');
DataRowNum=getappdata(0,'Auto2_4_2_DataRowNum');
COLOR_INDEX=getappdata(0,'Auto2_4_2_COLOR_INDEX');
Fileimname=getappdata(0,'Auto2_4_2_Fileimname');
ZIHAO_TU_YULAN=10;
CHOOSE=get(handles.listbox1,'Value');                %listbox的值
j=ceil(CHOOSE/length(FigureIndex));                  %获取所选标签位于MP第几行数据
CruveNum=get(handles.popupmenu4,'value')+2; 

b=mod(CHOOSE,length(FigureIndex));       %获取所选标签对应的数据点在FigureIndex的第几个Cell
if b==0
    b=length(FigureIndex);
end
for i=1:length(FigureIndex{b})
plot(handles.axes1,MP{j,FigureIndex{b}(1,i)}(:,1),MP{j,FigureIndex{b}(1,i)}(:,2),'linewidth',2,'color',COLOR_INDEX(i,:));
hold on
Outstr{1,i}=[Fileimname{j,FigureIndex{b}(1,i)},':',num2str(MaxKraft(j,FigureIndex{b}(1,i)),'%.2f'),'N'];
Legendtext{1,i}=Fileimname{j,FigureIndex{b}(1,i)};
MaxY(i)=MaxKraft(j,FigureIndex{b}(1,i));
end
axis(handles.axes1,[0 inf,0,max(MaxY)*1.1])
grid on;
xlabel(handles.axes1,'Weg/位移[mm]','FontSize',ZIHAO_TU_YULAN);
ylabel(handles.axes1,'Kraft/力[N]','FontSize',ZIHAO_TU_YULAN);
legend(handles.axes1,Legendtext,'location','northwest','Interpreter','none')
set(handles.text9,'String',Outstr)


% --- Executes during object creation, after setting all properties.
function listbox1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)

% --- Executes during object creation, after setting all properties.
function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Fahrzeugcode (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes on selection change in popupmenu2.
function popupmenu2_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function popupmenu2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu3.
function popupmenu3_Callback(hObject, eventdata, handles)

% --- Executes during object creation, after setting all properties.
function popupmenu3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
clear MP
DATA_TYPE_KRAFT=get(handles.popupmenu3,'value');      %读取数据第几列为力
DATA_TYPE_WEG=get(handles.popupmenu2,'value');          %读取数据第几列为位移
CruveNum=get(handles.popupmenu4,'value')+2;             %读取一张图几条曲线
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据','MultiSelect','on');
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
    return;
else
    Datatotalnum=length(filename);
    if mod(Datatotalnum,14)==0             %检验数据是否为14的整数倍
        DataRowNum=Datatotalnum/14;
        t1=waitbar(0,'正在读入数据');
        for j=1:DataRowNum
            for i=1:14
                Fileindex=i+(j-1)*14;
                k = find('.'==filename{Fileindex});
                Fileimname{j,i}= filename{Fileindex}(1:k-1);
                Filename{j,i}=strcat(pathname,filename{Fileindex});
                [Type Sheet Format]=xlsfinfo(Filename{j,i}) ;
                sheet{Fileindex}=Sheet;
                %自检验数据为Sheet几
                z=1;                
                while z<= length(Sheet) 
                    a=xlsread(Filename{j,i},char(sheet{1,Fileindex}(1,z)));
                    b=size(a);
                    if b>1
                        break
                    else
                        z=z+1;
                        continue
                    end
                end
                MP_MITTLE{j,i}=a;
                %%%%%%%%%%%%%%%%%%%%%%
                MP{j,i}(:,1)=MP_MITTLE{j,i}(:,DATA_TYPE_WEG);
                MP{j,i}(:,2)=MP_MITTLE{j,i}(:,DATA_TYPE_KRAFT);
                MaxKraft(j,i)=max(MP{j,i}(:,2));
                waitbar(Fileindex/length(filename));
                try
                    system('taskkill/IM excel.exe');
                end
            end
        end
    else
        msgbox('数据数量错误，应为14的倍数');
        return;
    end
end
close(t1);
if CruveNum~=14
    PictureNumEveryData=ceil(14/CruveNum);           %每组数据有几幅图
    for i=1:PictureNumEveryData-1
        for j=1:CruveNum
            FigureIndex{i}(j)=j+(i-1)*CruveNum;               %将图序号按照figure总数进行切割，方便后续调用
        end
    end
    z= FigureIndex{i}(j)+1;
    k=1;
    while z<=14
        FigureIndex{i+1}(k)=z;
        k=k+1;
        z=z+1;
    end 
else
    FigureIndex{1}=[1,2,3,4,5,6,7,8,9,10,11,12,13,14];    
end
   k=1;
    for i=1:DataRowNum
        for j=1:length(FigureIndex)
            ListText{k}=['Teil',num2str(i),'-',num2str(j)];                 %list标签
            k=k+1;
        end
    end
setappdata(0,'Auto2_4_2_MP',MP);
setappdata(0,'Auto2_4_2_pathname',pathname);
setappdata(0,'Auto2_4_2_filename',filename);
setappdata(0,'Auto2_4_2_FigureIndex',FigureIndex);
setappdata(0,'Auto2_4_2_MaxKraft',MaxKraft);
setappdata(0,'Auto2_4_2_DataRowNum',DataRowNum);
setappdata(0,'Auto2_4_2_Fileimname',Fileimname);
set(handles.listbox1,'String',ListText);
set(handles.listbox1,'Value',1);
set(handles.pushbutton2,'enable','on')
msgbox('数据导入成功');



% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
MP=getappdata(0,'Auto2_4_2_MP');
FigureIndex=getappdata(0,'Auto2_4_2_FigureIndex');
MaxKraft=getappdata(0,'Auto2_4_2_MaxKraft');
DataRowNum=getappdata(0,'Auto2_4_2_DataRowNum');
COLOR_INDEX=getappdata(0,'Auto2_4_2_COLOR_INDEX');
Fileimname=getappdata(0,'Auto2_4_2_Fileimname');
pathname=getappdata(0,'Auto2_4_2_pathname');
filename=getappdata(0,'Auto2_4_2_filename');
ZIHAO_TU_YULAN=20;
Fileadress=strcat(pathname,'result\');
if ~exist('pathname\result','dir')
    mkdir(pathname,'result');
end

t1=waitbar(0,'正在生成图片') ;
CHOOSELength= length(get(handles.listbox1,'String'));
CruveNum=get(handles.popupmenu4,'value')+2;
for k=1:CHOOSELength
    h=figure(k);
    set(h,'visible','off');
    set(h,'color','w')
    CHOOSE=k;                %listbox的值
    j=ceil(CHOOSE/length(FigureIndex));                  %获取所选标签位于MP第几行数据
    b=mod(CHOOSE,length(FigureIndex));       %获取所选标签对应的数据点在FigureIndex的第几个Cell
    if b==0
        b=length(FigureIndex);
    end
    for i=1:length(FigureIndex{b})
        plot(MP{j,FigureIndex{b}(1,i)}(:,1),MP{j,FigureIndex{b}(1,i)}(:,2),'linewidth',2,'color',COLOR_INDEX(i,:));
        hold on      
        Legendtext{1,i}=Fileimname{j,FigureIndex{b}(1,i)};
        MaxY(i)=MaxKraft(j,FigureIndex{b}(1,i));
    end
    set(h,'position',[100,100,1300,800]);
    hold off
    axis([0 inf,0,max(MaxY)*1.1])
    grid on;
    xlabel('Weg/位移[mm]','FontSize',ZIHAO_TU_YULAN);
    ylabel('Kraft/力[N]','FontSize',ZIHAO_TU_YULAN);
    legend(Legendtext,'location','northwest','Interpreter','none');
    set(gca,'FontSize',ZIHAO_TU_YULAN);
    sfilename=[Fileadress,num2str(k),'.jpg'];
    f=getframe(h);
    imwrite(f.cdata,sfilename);
    waitbar(k/CHOOSELength);
    close(h);
end
close(t1)
t2=waitbar(0,'正在生成报告');
biaotihao=10;
filespec_user=[Fileadress,strcat(num2str(year(now)),num2str(month(now)),num2str(day(now)),num2str(hour(now)),num2str(second(now),'%.0f')),'.doc'];
try 
Word=actxGetRunningServer('Word.Application');
catch 
Word=actxserver('Word.Application'); 
end
Word.Visible =0; % 使word为可见；或set(Word, 'Visible', 1); 
%===打开word文件，如果路径下没有则创建一个空白文档打开========================
if exist(filespec_user,'file')
Document=Word.Documents.Open(filespec_user);
else
Document=Word.Documents.Add;
Document.SaveAs2(filespec_user);
end
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;
waitbar(0.1)
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;
a=get(handles.popupmenu6,'String');
b=get(handles.popupmenu6,'Value');
headline=strcat('试验结果--',a{b});
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=biaotihao; % 字体大小
Content.Font.NameAscii='Arial';
Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落         
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落
Tab1 = Document.Tables.Add(Selection.Range, DataRowNum*2+2,9);
 DTI = Document.Tables.Item(1); % 表格句柄
 DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
 DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
 lc=28.381133333333333333333333333333; %厘米换算
 DTI.Columns.Item(1).Width =3*lc;
 for i = 2:8
     DTI.Columns.Item(i).Width = 1.5*lc;
 end
 DTI.Columns.Item(9).Width = 1.7*lc;
 DTI.Range.Paragraphs.Alignment='wdAlignParagraphCenter';
 DTI.Range.Font.NameAscii='Arial';
 DTI.Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
 DTI.Cell(1,1).Range.Text ='序号 Nr.';
 DTI.Cell(DataRowNum+2,1).Range.Text ='序号 Nr.';
 DTI.Cell(1,9).Range.Text ='Sollwert 理论值';
 DTI.Cell(2,9).Merge(DTI.Cell(DataRowNum*2+2,9));
 waitbar(0.2)
 CHECKKRAFT=get(handles.popupmenu6,'value');               %判断是安装力还是拆卸力

 if CHECKKRAFT==1
     DTI.Cell(2,9).Range.Text='<160 N';
 elseif  CHECKKRAFT==2   
     DTI.Cell(2,9).Range.Text='<70-200 N';
 end
%写入标题
 for k=1:7
     DTI.Cell(1,k+1).Range.Text = num2str(k);   
 end
 for k=8:14
     DTI.Cell(DataRowNum+2,k-6).Range.Text = num2str(k);
 end 
 %写入单元格颜色
  for k=1:9
   DTI.Cell(1,k).Range.Shading.BackgroundPatternColor=14857101;
  end
  for k=1:8
      DTI.Cell(DataRowNum+2,k).Range.Shading.BackgroundPatternColor=14857101;
  end
   waitbar(0.5)
 %写入数据
 for i=1:DataRowNum
     DTI.Cell(i+1,1).Range.Text =Fileimname{i,1};
     for k=1:7
         DTI.Cell(i+1,k+1).Range.Text = num2str(MaxKraft(i,k),'%.2f');
         if CHECKKRAFT==1
             if MaxKraft(i,k)>=160
                 DTI.Cell(i+1,k+1).Range.Font.Colorindex='wdRed';
                 DTI.Cell(i+1,k+1).Range.Font.Bold=1;
             end
         elseif CHECKKRAFT==2
             if MaxKraft(i,k)<70 || MaxKraft(i,k)>200
                 DTI.Cell(i+1,k+1).Range.Font.Colorindex='wdRed';
                 DTI.Cell(i+1,k+1).Range.Font.Bold=1;
             end
         end
     end
 end
 
  for i=1:DataRowNum
     DTI.Cell(DataRowNum+1+i+1,1).Range.Text =Fileimname{i,1};
     for k=1:7
         DTI.Cell(DataRowNum+1+i+1,k+1).Range.Text = num2str(MaxKraft(i,k+7),'%.2f');
         if CHECKKRAFT==1
             if MaxKraft(i,k+7)>=160
                 DTI.Cell(DataRowNum+1+i+1,k+1).Range.Font.Colorindex='wdRed';
                 DTI.Cell(DataRowNum+1+i+1,k+1).Range.Font.Bold=1;
             end
         elseif CHECKKRAFT==2
             if MaxKraft(i,k+7)<70 || MaxKraft(i,k+7)>200
                 DTI.Cell(DataRowNum+1+i+1,k+1).Range.Font.Colorindex='wdRed';
                 DTI.Cell(DataRowNum+1+i+1,k+1).Range.Font.Bold=1;
             end
         end
     end
  end
  Selection.Start = Content.end;
  Selection.TypeParagraph;
  InlineShapes=Document.InlineShapes;
  waitbar(0.7)
  for k=1:CHOOSELength
      sfilename1=[Fileadress,num2str(k),'.jpg'];
      handle=Selection.InlineShapes.AddPicture(sfilename1);
      Selection.Start = Content.end;
      Selection.TypeParagraph;
      list=get(handles.listbox1,'String');    
      headline=list{k};
      Selection.Text=headline;
      Selection.Font.NameAscii='Arial';
      Selection.Font.Size=biaotihao; % 字体大小
      Selection.Start = Content.end;
      Selection.TypeParagraph;
      Selection.Start = Content.end;
      Selection.TypeParagraph;
  end
  waitbar(0.9)
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
   %%%%%%%%%%%%输出报告生成信息到公共空间%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='VWi导水槽安装力试验';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
  Word.Visible =1;
  close(t2)
  
  
% --- Executes on selection change in popupmenu4.
function popupmenu4_Callback(hObject, eventdata, handles)
CruveNum=get(handles.popupmenu4,'value')+2;             %读取一张图几条曲线
DataRowNum=getappdata(0,'Auto2_4_2_DataRowNum');
if CruveNum~=14
    PictureNumEveryData=ceil(14/CruveNum);           %每组数据有几幅图
    for i=1:PictureNumEveryData-1
        for j=1:CruveNum
            FigureIndex{i}(j)=j+(i-1)*CruveNum;
        end
    end
    z= FigureIndex{i}(j)+1;
    k=1;
    while z<=14
        FigureIndex{i+1}(k)=z;
        k=k+1;
        z=z+1;
    end   
else
    FigureIndex{1}=[1,2,3,4,5,6,7,8,9,10,11,12,13,14];
end
 k=1;
    for i=1:DataRowNum
        for j=1:length(FigureIndex)
            ListText{k}=['Teil',num2str(i),'-',num2str(j)];
            k=k+1;
        end
    end
setappdata(0,'Auto2_4_2_FigureIndex',FigureIndex);
setappdata(0,'Auto2_4_2_DataRowNum', DataRowNum);
set(handles.listbox1,'String',ListText);
set(handles.listbox1,'Value',1);


% --- Executes during object creation, after setting all properties.
function popupmenu4_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu4 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu6.
function popupmenu6_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu6 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu6


% --- Executes during object creation, after setting all properties.
function popupmenu6_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu6 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
