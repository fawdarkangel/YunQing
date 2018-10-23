function varargout = Auto7_4_Configuration(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto7_4_Configuration_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto7_4_Configuration_OutputFcn, ...
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


% --- Executes just before Auto7_4_Configuration is made visible.
function Auto7_4_Configuration_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
load([cd,'\interface\Config\Auto7_4_Version.mat'])            %读取配置文件
DATA_TYPE=CONFIG_7_4.DATA_TYPE;                       %读取数据类型
DATA_TYPE_KRAFT=CONFIG_7_4.DATA_TYPE_KRAFT;      %读取数据第几列为力
DATA_TYPE_WEG=CONFIG_7_4.DATA_TYPE_WEG;          %读取数据第几列为位移
X_LABLE=CONFIG_7_4.X_LABLE;                                 %读取横坐标
Y_LABLE=CONFIG_7_4.Y_LABLE;                                 %读取纵坐标



setappdata(0,'CONFIG_7_4',CONFIG_7_4);                             %将配置写入内存
set(handles.popupmenu1,'Value',DATA_TYPE);       %设置数据类型选项框
set(handles.popupmenu2,'Value',DATA_TYPE_KRAFT);       %设置数据类型选项框
set(handles.popupmenu3,'Value',DATA_TYPE_WEG);       %设置数据类型选项框
set(handles.edit1,'String',X_LABLE);                       %写入输入框横坐标
set(handles.edit2,'String',Y_LABLE);                       %写入输入框纵坐标



handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto7_4_Configuration wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto7_4_Configuration_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)

CONFIG_7_4=getappdata(0,'CONFIG_7_4');                             %读取内存配置


DATA_TYPE=get(handles.popupmenu1,'Value');          %读取数据类型
DATA_TYPE_KRAFT=get(handles.popupmenu2,'Value');    %读取第几列为力
DATA_TYPE_WEG=get(handles.popupmenu3,'Value');     %读取第几列为位移
X_LABLE=get(handles.edit1,'String');                                       %读取横坐标标题
Y_LABLE=get(handles.edit2,'String');                                       %读取纵坐标标题



%%%%%%%%%%%%%%%%%%%%更新内存中配置信息%%%%%%%%%%%%%%%%%%
CONFIG_7_4.DATA_TYPE=DATA_TYPE;
CONFIG_7_4.DATA_TYPE_KRAFT=DATA_TYPE_KRAFT;
CONFIG_7_4.DATA_TYPE_WEG=DATA_TYPE_WEG;
CONFIG_7_4.X_LABLE=X_LABLE;
CONFIG_7_4.Y_LABLE=Y_LABLE;


setappdata(0,'CONFIG_7_4',CONFIG_7_4);                             %将配置写入内存

save([cd,'\interface\Config\Auto7_4_Version.mat'],'CONFIG_7_4');
msgbox('配置保存成功');


% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)

function popupmenu1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu2.
function popupmenu2_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function popupmenu2_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu3.
function popupmenu3_Callback(hObject, eventdata, handles)

function popupmenu3_CreateFcn(hObject, eventdata, handles)


if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



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
