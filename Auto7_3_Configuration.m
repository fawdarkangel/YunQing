function varargout = Auto7_3_Configuration(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto7_3_Configuration_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto7_3_Configuration_OutputFcn, ...
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


% --- Executes just before Auto7_3_Configuration is made visible.
function Auto7_3_Configuration_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
load([cd,'\model\Auto7_3_Version.mat'])            %读取配置文件
DATA_TYPE=CRUVE.DATA_TYPE;                       %读取数据类型
X_LABLE=CRUVE.X_LABLE;                                 %读取横坐标
Y_LABLE=CRUVE.Y_LABLE;                                 %读取纵坐标
TITLE_INDEX=CRUVE.TITLE_INDEX;                    %读取图片标题
Y_LABLE_DECIMALDIGITS=CRUVE.Y_LABLE_DECIMALDIGITS; %读取纵坐标小数位数
X_LABLE_DECIMALDIGITS=CRUVE.X_LABLE_DECIMALDIGITS; %读取纵坐标小数位数
LINECOLOR=CRUVE.LINECOLOR;                      %读取线条颜色
WIDTH=CRUVE.WIDTH;                                     %读取线宽
FONTSIZE=CRUVE.FONTSIZE;                            %读取字号
MAXKRAFT_INSERT= CRUVE.MAXKRAFT_INSERT;   %读取是否在图中标注最大值，1：标注，0：不标注
ASSISSTANT_LINE_CHECK=CRUVE.ASSISSTANT_LINE_CHECK;          %读取是否需要辅助线
ASSISSTANT_LINE_KRAFT=CRUVE.ASSISSTANT_LINE_KRAFT;           %读取辅助线大小
ASSISSTANT_LINE_COLORINDEX=CRUVE.ASSISSTANT_LINE_COLORINDEX;      %读取辅助线颜色
GRID_WIDTH=CRUVE.GRID_WIDTH;                     %获取网格密度 0：不加密 1：加密
DATA_TYPE_KRAFT=CRUVE.DATA_TYPE_KRAFT;      %读取数据第几列为力
DATA_TYPE_WEG=CRUVE.DATA_TYPE_WEG;          %读取数据第几列为位移
TITLEFONTSIZ=CRUVE.TITLEFONTSIZE;                 %读取标题字号
DATASHEET=CRUVE.DATASHEET;                             %读取数据位于Sheet几

%如果为Zwick激活popupmenu9，第几个Sheet为数据
if DATA_TYPE==1
    set(handles.popupmenu9,'Enable','on');
else
    set(handles.popupmenu9,'Enable','off');
end

if TITLE_INDEX==7                                             %如果标题上次是自定义的话，初始化打开时即释放自定义按钮
 set(handles.pushbutton2,'Enable','on');
else
  set(handles.pushbutton2,'Enable','off');
end

setappdata(0,'CRUVE',CRUVE);                             %将配置写入内存
setappdata(0,'STAND_TITLE',STAND_TITLE);        %将标题写入内存

set(handles.edit1,'String',X_LABLE);                       %写入输入框横坐标
set(handles.edit2,'String',Y_LABLE);                       %写入输入框纵坐标
set(handles.popupmenu2,'Value',DATA_TYPE);       %设置数据类型选项框
set(handles.popupmenu1,'Value',TITLE_INDEX);     %设置图标标题选项框
set(handles.popupmenu3,'Value',Y_LABLE_DECIMALDIGITS);     %设置纵坐标小数位数
set(handles.popupmenu4,'Value',X_LABLE_DECIMALDIGITS);     %设置纵坐标小数位数
set(handles.popupmenu5,'Value',LINECOLOR);       %设置线条颜色
set(handles.edit3,'String',WIDTH);                           %写入线宽
set(handles.edit4,'String',FONTSIZE);                       %写入字号
set(handles.checkbox1,'Value',MAXKRAFT_INSERT)  %写入是否标注最大力值
set(handles.checkbox2,'Value',ASSISSTANT_LINE_CHECK);
set(handles.checkbox3,'Value',GRID_WIDTH);           %写入网格密度复选框
set(handles.popupmenu7,'Value',DATA_TYPE_KRAFT);       %写入数据第几列为力
set(handles.popupmenu8,'Value',DATA_TYPE_WEG);       %写入数据第几列为位移
set(handles.edit6,'String',TITLEFONTSIZ);       %写入数据第几列为位移
set(handles.popupmenu9,'Value',DATASHEET);   %写入第几个Sheet为数据


if ASSISSTANT_LINE_CHECK==1                                        %需要辅助横线
    set(handles.edit5,'Enable','on');
    set(handles.popupmenu6,'Enable','on');    
    set(handles.edit5,'String',ASSISSTANT_LINE_KRAFT);
    set(handles.popupmenu6,'Value',ASSISSTANT_LINE_COLORINDEX);
else                                                                                       %不需要辅助横线
    set(handles.edit5,'Enable','off');
    set(handles.popupmenu6,'Enable','off');
    set(handles.edit5,'String',ASSISSTANT_LINE_KRAFT);
    set(handles.popupmenu6,'Value',ASSISSTANT_LINE_COLORINDEX);
end



handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto7_3_Configuration wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto7_3_Configuration_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- 保存配置按钮.
function pushbutton1_Callback(hObject, eventdata, handles)

CRUVE=getappdata(0,'CRUVE');                             %读取内存中配置文件
STAND_TITLE=getappdata(0,'STAND_TITLE');           %读取内存中标题文件
DATA_TYPE=get(handles.popupmenu2,'Value');          %读取数据类型
TITLE_INDEX=get(handles.popupmenu1,'Value');         %读取图标标题索引

X_LABLE=get(handles.edit1,'String');                                       %读取横坐标标题
Y_LABLE=get(handles.edit2,'String');                                       %读取纵坐标标题
Y_LABLE_DECIMALDIGITS=get(handles.popupmenu3,'Value'); %读取纵坐标小数位数
X_LABLE_DECIMALDIGITS=get(handles.popupmenu4,'Value'); %读取横坐标小数位数
LINECOLOR=get(handles.popupmenu5,'Value');                     %读取线条颜色
WIDTH=get(handles.edit3,'String');                                         %读取线宽
FONTSIZE=get(handles.edit4,'String');                                     %读取字号
MAXKRAFT_INSERT= get(handles.checkbox1,'Value');   %读取是否在图中标注最大值，1：标注，0：不标注
ASSISSTANT_LINE_CHECK=get(handles.checkbox2,'Value');  %读取是否需要辅助线 1：需要  0：不需要 
ASSISSTANT_LINE_KRAFT=get(handles.edit5,'String');           %读取辅助线力值
ASSISSTANT_LINE_COLORINDEX=get(handles.popupmenu6,'Value');  %读取辅助线颜色序号
GRID_WIDTH=get(handles.checkbox3,'Value');  %读取是网格密度 1：加密  0：不加密; 
DATA_TYPE_KRAFT=get(handles.popupmenu7,'Value');    %读取第几列为力
DATA_TYPE_WEG=get(handles.popupmenu8,'Value');     %读取第几列为位移
TITLEFONTSIZE=get(handles.edit6,'String');       %读取标题字号
DATASHEET=get(handles.popupmenu9,'Value');   %写入第几个Sheet为数据

%%%%%%%%%%%选择图标标题%%%%%%%%%%%%%%%
switch TITLE_INDEX
    case 1
         for i=1:1000
            STAND_TITLE{i}=[' '];
        end
    case 2
        for i=1:1000
            STAND_TITLE{i}=['MP ',num2str(i)];
        end
    case 3
        for i=1:1000
            STAND_TITLE{i}=['MP ',num2str(i),'#'];
        end
    case 4
        for i=1:1000
            STAND_TITLE{i}=['Teil ',num2str(i)];
        end
    case 5
        for i=1:1000
            STAND_TITLE{i}=['Teil ',num2str(i),'#'];
        end
    case 6
        for i=1:1000
            STAND_TITLE{i}=[num2str(i),'#'];
        end
    case 7        
        set(handles.pushbutton1,'Enable','on');
        STAND_TITLE=getappdata(0,'STAND_TITLE');
    case 8
        filename=getappdata(0,'filename');
        if isempty(filename)
            msgbox('未检测到文件名，请先导入数据');
            return
        end
        for i=1:length(filename)
            n(i,1)=find('.'==filename{1,i});
            STAND_TITLE{i}=filename{1,i}(1:n(i,1)-1);
        end        
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%更新配置文件%%%%%%%%%%%%%%%%%%
CRUVE.DATA_TYPE=DATA_TYPE;
CRUVE.X_LABLE=X_LABLE;
CRUVE.Y_LABLE=Y_LABLE;
CRUVE.TITLE_INDEX=TITLE_INDEX;
CRUVE.Y_LABLE_DECIMALDIGITS=Y_LABLE_DECIMALDIGITS;
CRUVE.X_LABLE_DECIMALDIGITS=X_LABLE_DECIMALDIGITS;
CRUVE.LINECOLOR=LINECOLOR;
CRUVE.WIDTH=WIDTH;
CRUVE.FONTSIZE=FONTSIZE;
CRUVE.MAXKRAFT_INSERT=MAXKRAFT_INSERT;
CRUVE.ASSISSTANT_LINE_CHECK=ASSISSTANT_LINE_CHECK;        
CRUVE.ASSISSTANT_LINE_KRAFT=ASSISSTANT_LINE_KRAFT;        
CRUVE.ASSISSTANT_LINE_COLORINDEX=ASSISSTANT_LINE_COLORINDEX;     
CRUVE.GRID_WIDTH=GRID_WIDTH;
CRUVE.DATA_TYPE_KRAFT=DATA_TYPE_KRAFT;
CRUVE.DATA_TYPE_WEG=DATA_TYPE_WEG;
CRUVE.TITLEFONTSIZE=TITLEFONTSIZE;
CRUVE.DATASHEET=DATASHEET;
setappdata(0,'CRUVE',CRUVE);
setappdata(0,'STAND_TITLE',STAND_TITLE);
if TITLE_INDEX==7&&isempty(STAND_TITLE)
    msgbox('请导入自定义标题EXCEL')
    return
end

save([cd,'\model\Auto7_3_Version.mat'],'CRUVE','STAND_TITLE')
msgbox('配置保存成功');

function edit1_Callback(hObject, eventdata, handles)

function edit1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit2_Callback(hObject, eventdata, handles)

function edit2_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)
TITLE_INDEX=get(handles.popupmenu1,'Value');
if TITLE_INDEX==7
 set(handles.pushbutton2,'Enable','on');
else
  set(handles.pushbutton2,'Enable','off');
end


function popupmenu1_CreateFcn(hObject, eventdata, handles)





if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据');
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
    return;
else
    Filename=strcat(pathname,filename);
    [Type Sheet Format]=xlsfinfo(Filename) ;
    sheet=Sheet;
    [NUM ROW STAND_TITLE]=xlsread(Filename,char(sheet(1,1)));    
    setappdata(0,'STAND_TITLE',STAND_TITLE);
    
    
    msgbox('标题导入成功');
end


function popupmenu2_Callback(hObject, eventdata, handles)

DATA_TYPE=get(handles.popupmenu2,'Value');          %读取数据类型
if DATA_TYPE==1
    set(handles.popupmenu9,'Enable','on');
else
    set(handles.popupmenu9,'Enable','off');
end
    

function popupmenu2_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu3.
function popupmenu3_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function popupmenu3_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu4.
function popupmenu4_Callback(hObject, eventdata, handles)

function popupmenu4_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu5.
function popupmenu5_Callback(hObject, eventdata, handles)

function popupmenu5_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit3_Callback(hObject, eventdata, handles)

function edit3_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit4_Callback(hObject, eventdata, handles)

function edit4_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in checkbox1.
function checkbox1_Callback(hObject, eventdata, handles)


% --- Executes on button press in checkbox2.
function checkbox2_Callback(hObject, eventdata, handles)
ASSISSTANT_LINE_CHECK=get(handles.checkbox2,'Value');
if ASSISSTANT_LINE_CHECK==1
    set(handles.edit5,'Enable','on');
    set(handles.popupmenu6,'Enable','on');
else
    set(handles.edit5,'Enable','off');
    set(handles.popupmenu6,'Enable','off');
end


function edit5_Callback(hObject, eventdata, handles)

function edit5_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu6.
function popupmenu6_Callback(hObject, eventdata, handles)

function popupmenu6_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in checkbox3.
function checkbox3_Callback(hObject, eventdata, handles)


% --- Executes on selection change in popupmenu7.
function popupmenu7_Callback(hObject, eventdata, handles)

function popupmenu7_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu8.
function popupmenu8_Callback(hObject, eventdata, handles)

function popupmenu8_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit6_Callback(hObject, eventdata, handles)

function edit6_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu9.
function popupmenu9_Callback(hObject, eventdata, handles)

function popupmenu9_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
