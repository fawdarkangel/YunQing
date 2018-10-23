function varargout = Auto7_1(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto7_1_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto7_1_OutputFcn, ...
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


% --- Executes just before Auto7_1 is made visible.
function Auto7_1_OpeningFcn(hObject, eventdata, handles, varargin)

handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto7_1 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto7_1_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;

function edit1_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function popupmenu1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
function UNITPLOT(MP_X,MP_Y,zihao,Y_UNIT,TITLE,hObject, eventdata, handles)

zihao=zihao;
Y_UNIT=Y_UNIT;
MP(:,1)=MP_X;
MP(:,2)=MP_Y;
switch Y_UNIT
    case 1

 plot(MP(:,1),MP(:,2),'linewidth',2);
    a=max(MP(:,1))*1.1; b=max(MP(:,2))*1.1;
    axis([0 a 0 b]);
    xlabel('Weg(mm)','FontSize',zihao);ylabel('Kraft(N)','FontSize',zihao);
   
    case 2
         plot(MP(:,1),MP(:,2)/1000,'linewidth',2);
    a=max(MP(:,1))*1.1; b=max(MP(:,2))*1.1/1000;
    axis([0 a 0 b]);
    xlabel('Weg(mm)','FontSize',zihao);ylabel('Kraft(KN)','FontSize',zihao);
        
        end
     title(TITLE,'FontSize',zihao);
    grid on;
    

    

% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);

global newpath

oldpath=cd;
if isempty(newpath)|| ~exist('newpath')
     newpath=cd;
 end
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据',newpath);

if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
    return;
else
    TITLE=get(handles.edit1,'String');
    list=get(handles.popupmenu1,'String');
val1=get(handles.popupmenu1,'Value');
zihao=str2num(get(handles.edit2,'String'));
 str=strcat(pathname,filename);
    newpath=pathname; 
end
Y_UNIT=get(handles.popupmenu2,'Value');

switch val1               %选择数据类型
%%%%%%%%%%%%%%%%%%%%% Zwick模块%%%%%%%%%%%%%%%%%5
    case 1                      %case 1 zwick数据类型
 
[Type Sheet Format]=xlsfinfo(str)
MP=xlsread(str,Sheet{4}); 
clear global h;
global h;
    h=figure;
    set(h,'visible','off');
    UNITPLOT(MP(:,1),MP(:,2),zihao,Y_UNIT,TITLE)
     
cla(handles.axes1);   
    switch Y_UNIT
         case 1
     plot(handles.axes1,MP(:,1),MP(:,2),'linewidth',2);
    a=max(MP(:,1))*1.1; b=max(MP(:,2))*1.1;
    axis(handles.axes1,[0 a 0 b]);
    xlabel(handles.axes1,'Weg(mm)');ylabel(handles.axes1,'Kraft(N)');
      case 2
       plot(handles.axes1,MP(:,1),MP(:,2)/1000,'linewidth',2);
    a=max(MP(:,1))*1.1; b=max(MP(:,2))*1.1/1000;
    axis(handles.axes1,[0 a 0 b]);
    xlabel(handles.axes1,'Weg(mm)');ylabel(handles.axes1,'Kraft(KN)');
    end
    title(handles.axes1,get(handles.edit1,'String'));
    grid on;
    msgbox('图片生成完毕，请保存图像');
   set(handles.pushbutton2,'Enable','on');
    set(handles.edit3,'String',num2str(b/1.1,'%.2f'));     %输出最大力
    set(handles.edit4,'String',num2str(a/1.1,'%.2f'));     %输出最大变形
end
% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
global h;
handles=guihandles;
guidata(hObject,handles);
if ~isempty(h)
[filename1,pathname1]=uiputfile({'*.jpg','JPG files';'*.bmp','BMP files'},'保存');%输出函数
if isequal(filename1,0)||isequal(pathname1,0)
    return
else
end
sa=strcat(pathname1,filename1);
saveas(h,sa);
close(h);
set(handles.pushbutton2,'Enable','off');
end


% --- Executes on selection change in popupmenu2.
function popupmenu2_Callback(hObject, eventdata, handles)

function popupmenu2_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit2_Callback(hObject, eventdata, handles)

function edit2_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit4_Callback(hObject, eventdata, handles)

function edit4_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit3_Callback(hObject, eventdata, handles)

function edit3_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
