function varargout = Auto2_3_Config(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto2_3_Config_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto2_3_Config_OutputFcn, ...
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


% --- Executes just before Auto2_3_Config is made visible.
function Auto2_3_Config_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
load([cd,'\interface\Config\Auto2_3_Config.mat'])            %读取配置文件

set(handles.edit1,'String',num2str(CONFIG.FONTSIZE));                       %写入输入框横坐标
set(handles.edit2,'String',num2str(CONFIG.TITLEFONTSIZE));                       %写入输入框横坐标
set(handles.edit3,'String',num2str(CONFIG.Figure_Width));                       %写入输入框横坐标
set(handles.edit4,'String',num2str(CONFIG.Figure_Height));                       %写入输入框横坐标

setappdata(0,'AUTO_2_3CONFIG',CONFIG);

handles.output = hObject;

% Update handles structure
guidata(hObject, handles);



% --- Outputs from this function are returned to the command line.
function varargout = Auto2_3_Config_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;



function edit1_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit2_Callback(hObject, eventdata, handles)

function edit2_CreateFcn(hObject, eventdata, handles)

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


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
CONFIG=getappdata(0,'AUTO_2_3CONFIG');    

CONFIG.FONTSIZE=str2num(get(handles.edit1,'String'));                       %写入输入框横坐标
CONFIG.TITLEFONTSIZE=str2num(get(handles.edit2,'String'));    
CONFIG.Figure_Width=str2num(get(handles.edit3,'String'));    
CONFIG.Figure_Height=str2num(get(handles.edit4,'String'));    

setappdata(0,'AUTO_2_3CONFIG',CONFIG);
save([cd,'\interface\Config\Auto2_3_Config.mat'],'CONFIG');
msgbox('配置保存成功');
