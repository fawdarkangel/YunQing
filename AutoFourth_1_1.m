function varargout = AutoFourth_1_1(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @AutoFourth_1_1_OpeningFcn, ...
                   'gui_OutputFcn',  @AutoFourth_1_1_OutputFcn, ...
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


% --- Executes just before AutoFourth_1_1 is made visible.
function AutoFourth_1_1_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to AutoFourth_1_1 (see VARARGIN)

% Choose default command line output for AutoFourth_1_1
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes AutoFourth_1_1 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = AutoFourth_1_1_OutputFcn(hObject, eventdata, handles) 
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
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu1 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu1


% --- Executes during object creation, after setting all properties.
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
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','ѡ������',newpath);
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('�����ļ�ʧ��');
  return;
else
    zihao=20;
    newpath=pathname; 
    Filename=strcat(pathname,filename);
    [Type Sheet Format]=xlsfinfo(Filename)
    MP=xlsread(Filename,Sheet{1}); 
    b1=str2num(get(handles.edit2,'String'));
      b2=str2num(get(handles.edit3,'String'));
       if min(MP(:,2))<0
             MP(:,2)=MP(:,2).*(-1);
             MP(:,3)=MP(:,3).*(-1);
       end
       %%%������2�е����У�ȷ���ڶ���Ϊλ�ƣ�������Ϊ��
       MP_M(:,1)=MP(:,2);
       MP_M(:,2)=MP(:,3);
       MP(:,2)=MP_M(:,2);
       MP(:,3)=MP_M(:,1);
       
      MP_FINAL(:,1)=rmmissing(MP(:,1));
        MP_FINAL(:,2)=rmmissing(MP(:,2));
        MP_FINAL(:,3)=rmmissing(MP(:,3));
end
 for j=1:length(MP_FINAL)
MPmin=b1;
if MP_FINAL(j,2)>=MPmin
a2=MP_FINAL(j,2);Lmin=j;   %a2Ϊ���Իع�������ʼ��
break;
end
    end

 

    for j=1:length(MP_FINAL)
MPmax=b2;if MP_FINAL(j,2)>=MPmax
a3=MP_FINAL(j,2);Lmax=j;   %a2Ϊ���Իع�������ֹ��
break;
end
    end

MPx=MP_FINAL(Lmin:Lmax,1:3); %MPxΪ���Իع�����

  [p_1,p_2]=polyfit(MPx(:,3),MPx(:,2),1);%p1(1,1)Ϊб��
  p1=p_1(1,1);
Y1_INDEX=find(MP_FINAL(:,2)>35000,1); %��Ѱ��һ������35000N�ĵ�����
Y1=MP_FINAL(Y1_INDEX,2);%Y1Ϊ��һ������35000N����ֵ�����ڼ����һ���������յ�
X2=MP_FINAL(Y1_INDEX,3);%X2Ϊ��һ������35000N����ֵ��Ӧ�ĺ�����
 

%X1=Y1/p1;     %[X1 Y1]Ϊ��һ����������ֹ��
Y_NIHE=p1*MP_FINAL(:,3)+p_1(1,2);
Y_NIHE_INDEX=find(Y_NIHE>35000,1);
B=X2-Y1/p1; %�ڶ��������߽ؾ࣬�����Ա��ν��

%% Zwickģ��
global h;
    h=figure;
 set(h,'visible','off');
  
    set(h,'position',[100,100,1300,800]); 
    plot(MP_FINAL(:,3),MP_FINAL(:,2)./1000,'linewidth',2);
    hold on
    %plot([0,X1],[0,Y1]/1000,'--','linewidth',2,'Color','r'); %��һ��������
    plot(MP_FINAL(1:Y_NIHE_INDEX,3),Y_NIHE(1:Y_NIHE_INDEX)/1000,'--','linewidth',2,'Color','r'); %��һ��������
    plot([0,X2],[Y1,Y1]/1000,'--','linewidth',2,'Color','r');%35000N����
    plot([B X2],[0 Y1]./1000,'--','linewidth',2,'Color','r');
    axis([0 inf 0 inf]);grid on;
    grid minor;
      set(gca,'FontSize',zihao);
       text(B,1,[num2str(B,'%.2f'),'mm'],'FontSize',zihao);
       xlabel('Weg[mm]','FontSize',zihao);ylabel('Kraft[kN]','FontSize',zihao);
       title(get(handles.edit1,'String'),'FontSize',zihao);
     set(gca,'LineWid',2);
     
       AXES_ZIHAO=10;
       cla(handles.axes1);     
  plot(handles.axes1,MP_FINAL(:,3),MP_FINAL(:,2)./1000,'linewidth',2);
    hold on
    %plot(handles.axes1,[0,X1],[0,Y1]/1000,'--','linewidth',2,'Color','r'); %��һ��������
    plot(handles.axes1,MP_FINAL(1:Y_NIHE_INDEX,3),Y_NIHE(1:Y_NIHE_INDEX)/1000,'--','linewidth',2,'Color','r'); %��һ��������
    plot(handles.axes1,[0,X2],[Y1,Y1]/1000,'--','linewidth',2,'Color','r');%35000N����
    plot(handles.axes1,[B X2],[0 Y1]./1000,'--','linewidth',2,'Color','r');
    axis(handles.axes1,[0 inf 0 inf]);grid on
        text(handles.axes1,B,1,[num2str(B,'%.2f'),'mm'],'FontSize',AXES_ZIHAO);
       xlabel(handles.axes1,'Weg[mm]','FontSize',AXES_ZIHAO);ylabel(handles.axes1,'Kraft[kN]','FontSize',AXES_ZIHAO);
       title(handles.axes1,get(handles.edit1,'String'),'FontSize',AXES_ZIHAO);
        msgbox('ͼƬ������ϣ��뱣��ͼ��');
        set(handles.pushbutton2,'Enable','on');
% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
global h;
handles=guihandles;
guidata(hObject,handles);
if ~isempty(h)
[filename1,pathname1]=uiputfile({'*.jpg','JPG files';'*.bmp','BMP files'},'����');%�������
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
