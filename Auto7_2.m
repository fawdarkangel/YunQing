function varargout = Auto7_2(varargin)
% AUTO7_2 MATLAB code for Auto7_2.fig
%      AUTO7_2, by itself, creates a new AUTO7_2 or raises the existing
%      singleton*.
%
%      H = AUTO7_2 returns the handle to a new AUTO7_2 or the handle to
%      the existing singleton*.
%
%      AUTO7_2('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in AUTO7_2.M with the given input arguments.
%
%      AUTO7_2('Property','Value',...) creates a new AUTO7_2 or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before Auto7_2_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to Auto7_2_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help Auto7_2

% Last Modified by GUIDE v2.5 25-Apr-2018 08:13:02

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto7_2_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto7_2_OutputFcn, ...
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


% --- Executes just before Auto7_2 is made visible.
function Auto7_2_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to Auto7_2 (see VARARGIN)

% Choose default command line output for Auto7_2
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto7_2 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto7_2_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;

function edit1_Callback(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit1 as text
%        str2double(get(hObject,'String')) returns contents of edit1 as a double


% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)

function popupmenu1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


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
    zihao=20;
    newpath=pathname; 
    val1=get(handles.popupmenu2,'Value');
     Filename=strcat(pathname,filename);
       [Type Sheet Format]=xlsfinfo(Filename)
    switch val1
        case 1
         MP=xlsread(Filename,Sheet{4});    
         MP_FINAL(:,1)=MP(:,1);
         MP_FINAL(:,2)=MP(:,2);
       
            
            
    end
        
      b1=str2num(get(handles.edit2,'String'));
      b2=str2num(get(handles.edit3,'String'));
  
     
end
 for j=1:length(MP_FINAL)
MPmin=b1;
if MP_FINAL(j,2)>=MPmin
a2=MP_FINAL(j,2);Lmin=j;   %a2为线性回归数据起始点
break;
end
    end

 

    for j=1:length(MP_FINAL)
         MPmax=b2;
if MP_FINAL(j,2)>=MPmax
a3=MP_FINAL(j,2);Lmax=j;   %a2为线性回归数据终止点
break;
end
    end

MPx=MP_FINAL(Lmin:Lmax,1:2); %MPx为线性回归数据

  [p_1,p_2]=polyfit(MPx(:,1),MPx(:,2),1);%p1(1,1)为斜率
  p1=p_1(1,1);
YMAX=max(MP_FINAL(:,2)); %找寻第一个大于35000N的点坐标
X_INDEX=max(MP_FINAL(:,2))/p1;
set(handles.edit4,'String',num2str(p1,'%.1f'));

%% Zwick模块
global h;
    h=figure;
 set(h,'visible','off');
  
    set(h,'position',[100,100,1300,800]); 
    plot(MP_FINAL(:,1),MP_FINAL(:,2)./1000,'linewidth',2);
    hold on
    plot([0,X_INDEX],[0,YMAX]/1000,'--','linewidth',2,'Color','r'); %第一条辅助线
      axis([0 inf 0 inf]);grid on;
         set(gca,'FontSize',zihao);
             xlabel('Weg[mm]','FontSize',zihao);ylabel('Kraft[kN]','FontSize',zihao);
       title(get(handles.edit1,'String'),'FontSize',zihao);
     set(gca,'LineWid',2);
     
       AXES_ZIHAO=10;
       cla(handles.axes1);     
  plot(handles.axes1,MP_FINAL(:,1),MP_FINAL(:,2)./1000,'linewidth',2);
    hold on
    plot(handles.axes1,[0,X_INDEX],[0,YMAX]/1000,'--','linewidth',2,'Color','r'); %第一条辅助线

    axis(handles.axes1,[0 inf 0 inf]);grid on
    xlabel(handles.axes1,'Weg[mm]','FontSize',AXES_ZIHAO);ylabel(handles.axes1,'Kraft[kN]','FontSize',AXES_ZIHAO);
       title(handles.axes1,get(handles.edit1,'String'),'FontSize',AXES_ZIHAO);
        msgbox('图片生成完毕，请保存图像');
        set(handles.pushbutton2,'Enable','on');
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


% --- Executes on selection change in popupmenu2.
function popupmenu2_Callback(hObject, eventdata, handles)

function popupmenu2_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
