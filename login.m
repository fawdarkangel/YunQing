function varargout = login(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @login_OpeningFcn, ...
                   'gui_OutputFcn',  @login_OutputFcn, ...
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


% --- Executes just before login is made visible.
function login_OpeningFcn(hObject, eventdata, handles, varargin)

handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes login wait for user response (see UIRESUME)
% uiwait(handles.login);


% --- Outputs from this function are returned to the command line.
function varargout = login_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;



function edit1_Callback(hObject, eventdata, handles)

if get(gcf,'CurrentCharacter')==13
    pushbutton1_Callback(hObject, eventdata, handles)
    
end

% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)

%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit2_Callback(hObject, eventdata, handles)

if get(gcf,'CurrentCharacter')==13
    pushbutton1_Callback(hObject, eventdata, handles)
    
end

% --- Executes during object creation, after setting all properties.
function edit2_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
%STAND_VERSION=1.21;
t1=waitbar(0,'正在初始化网络……');
[status,cmdout]=dos('net group "domain admins" /domain'); %判断是否连接公司内网
waitbar(0.1);
if status==0
    C=load('\\faw-vw\fs\org\PE\T-E-VC-2-2\黄禹霆\12-数据处理平台\login.mat');
elseif status==2
    msgbox('请连接公司内网');
    return;
else
    errordlg('错误代码：1001','错误');
    return;
end
waitbar(0.5);
b1=get(handles.edit1,'String');
NAME=char(b1+16);
PASSWORD=char(get(handles.edit2,'String')+13);
NAME_INDEX=find(strcmp(C.login,NAME));
if isempty(NAME_INDEX)
    msgbox('用户名或密码错误');
    close(t1)
    return
else
    if strcmp(C.login{NAME_INDEX,2},{PASSWORD})==0
        msgbox('密码错误');
        close(t1)
        return
    else
close(t1);  
        
        close(login);
        run YunQing;
          % if C.VERSION>STAND_VERSION
    %msgbox('软件不是最新版，可能有不可预见的BUG，请前往公共空间考取最新版软件');
%end
        LOCAL_ADDRESS=java.net.InetAddress.getLocalHost;
        COMPUTER_IP=char(LOCAL_ADDRESS.getHostAddress);
        COMPUTER_NAME=char(LOCAL_ADDRESS.getHostName);
       Fileaddress=[ '\\faw-vw\fs\org\PE\T-E-VC-2-2\黄禹霆\12-数据处理平台\login information.xlsx'];
          [num text alldata]=xlsread(Fileaddress);
            SZ=size(alldata,1);%SZ为当前工作表行数
                       
            Azuobiao=strcat('A',num2str(SZ+1));
        OUTPUT_INFORMATION{1,1}=b1;
        OUTPUT_INFORMATION{1,2}=COMPUTER_NAME;
          OUTPUT_INFORMATION{1,3}=COMPUTER_IP;
          OUTPUT_INFORMATION{1,4}=datestr(now,'yyyy-mm-dd HH:MM:SS');
          try
         xlswrite([Fileaddress], OUTPUT_INFORMATION,'Sheet1',[Azuobiao]);
          end
    end
end

% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
close(gcf);

%function figure1_KeyPressFcn(hObject, eventdata, handles)
