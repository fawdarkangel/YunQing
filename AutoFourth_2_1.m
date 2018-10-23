function varargout = AutoFourth_2_1(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @AutoFourth_2_1_OpeningFcn, ...
                   'gui_OutputFcn',  @AutoFourth_2_1_OutputFcn, ...
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


% --- Executes just before AutoFourth_2_1 is made visible.
function AutoFourth_2_1_OpeningFcn(hObject, eventdata, handles, varargin)

handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes AutoFourth_2_1 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = AutoFourth_2_1_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;

function edit1_Callback(hObject, eventdata, handles)

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
clear global MP_FINAL POINT1_INDEX Y1_INDEX AXES_ZIHAO X0 X1 Y1 Y_CRUVE zihao
global newpath MP_FINAL POINT1_INDEX Y1_INDEX AXES_ZIHAO X0 X1 Y1 Y_CRUVE zihao

oldpath=cd;
if isempty(newpath)|| ~exist('newpath')
     newpath=cd;
 end
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','ѡ������',newpath);
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('�����ļ�ʧ��');
  return;
else
    
    t1=waitbar(0,'��������ͼƬ');
    zihao=20;
    newpath=pathname; 
    Filename=strcat(pathname,filename);
    [Type Sheet Format]=xlsfinfo(Filename)
    MP=xlsread(Filename,Sheet{1}); 
    b1=str2num(get(handles.edit2,'String'));
      b2=str2num(get(handles.edit3,'String'));
       if min(MP(:,2))<0
             MP(:,2)=MP(:,2).*(-1);                                         %���λ��Ϊ��ֵ����-1������һ����
             MP(:,3)=MP(:,3).*(-1);
                      end
      MP_FINAL(:,1)=rmmissing(MP(:,1));   %rmmisiising����ȥ��������NAN
        MP_FINAL(:,2)=rmmissing(MP(:,2));
        MP_FINAL(:,3)=rmmissing(MP(:,3));
end
waitbar(0.5);
for j=1:length(MP_FINAL)
MPmin=b1;                                                           %b1Ϊ�û�����������ֵ����
if MP_FINAL(j,3)>=MPmin
a2=MP_FINAL(j,3);Lmin=j;   %a2Ϊ���Իع�������ʼ��
break;
end
    end


 

    for j=1:length(MP_FINAL)
MPmax=b2;if MP_FINAL(j,3)>=MPmax                 %b2Ϊ�û�����������ֵ����
a3=MP_FINAL(j,3);Lmax=j;   %a2Ϊ���Իع�������ֹ��
break;
end
    end

MPx=MP_FINAL(Lmin:Lmax,1:3); %MPxΪ���Իع�����

  [p_1,p_2]=polyfit(MPx(:,2),MPx(:,3),1);%p1(1,1)Ϊб��
  p1=p_1(1,1);                                                                  
  B=p_1(1,2);                                                                             %BΪ���߽ؾ�
Y1_INDEX=find(MP_FINAL(:,3)==max(MP_FINAL(:,3)));              %��Ѱ�����ֵ
X_CRUVE=MP_FINAL(1:Y1_INDEX,2);                                        %λ�ƴӵ�һ�㵽��ֵ���㣬�������ʼ�뿪���ߵĵ�
Y_CRUVE=X_CRUVE.*p1+B;                                                      %Y_CRUVEΪ���ֱ�ߵ�Y����ֵ��������ԭ�����������ұ����

Y1=MP_FINAL(Y1_INDEX,3);                                                    %Y1Ϊ������ֵ�����ڼ����һ���������յ�
X1=(Y1-B)/p1;                                                                           %X1�����ߺ�����
X0=-B/p1;                                                                               %X0Ϊ��������X�ύ������
sliderValue = get(handles.slider1,'Value');                                 %��ȡ����������ֵ������Ѱ�ұ����
ang = int32(sliderValue); 
for i=1:length(Y_CRUVE)
POINT1_INDEX=find((Y_CRUVE(:,1)-MP_FINAL(1:Y1_INDEX,3))>ang,1);     %find���������ԭ���ߵĲ�ֵ����angʱ�ĵ㣬��Ϊ���߱��������
end

%% Zwickģ��

      
     
       AXES_ZIHAO=10;
       cla(handles.axes1);     
           plot(handles.axes1,MP_FINAL(:,2),MP_FINAL(:,3)./1000,'linewidth',2);
hold on
plot(handles.axes1,[X0,X1],[0,Y1]/1000,'--','linewidth',2,'Color','r'); %��һ��������
 %plot(handles.axes1,X_CRUVE,Y_CRUVE/1000)
plot(handles.axes1,MP_FINAL(POINT1_INDEX,2),MP_FINAL(POINT1_INDEX,3)/1000, 'o', 'markerfacecolor', [ 1, 0, 0 ])
plot(handles.axes1,[0,MP_FINAL(POINT1_INDEX,2)],[MP_FINAL(POINT1_INDEX,3)/1000,MP_FINAL(POINT1_INDEX,3)/1000],'--','linewidth',2,'Color','r')
plot(handles.axes1,MP_FINAL(Y1_INDEX,2),MP_FINAL(Y1_INDEX,3)/1000, 'o', 'markerfacecolor', [ 1, 0, 0 ])
plot(handles.axes1,[0,MP_FINAL(Y1_INDEX,2)],[MP_FINAL(Y1_INDEX,3)/1000,MP_FINAL(Y1_INDEX,3)/1000],'--','linewidth',2,'Color','r')

waitbar(0.8)
    axis(handles.axes1,[0 max(MP_FINAL(:,2))*1.1 0 max(MP_FINAL(:,3))/1000*1.1]);grid on
        text(handles.axes1,MP_FINAL(POINT1_INDEX,2)+1,MP_FINAL(POINT1_INDEX,3)/1000,[num2str(MP_FINAL(POINT1_INDEX,3),'%.0f'),'N'],'FontSize',AXES_ZIHAO);
        text(handles.axes1,MP_FINAL(Y1_INDEX,2),MP_FINAL(Y1_INDEX,3)/1000+2,[num2str(MP_FINAL(Y1_INDEX,3),'%.0f'),'N'],'FontSize',AXES_ZIHAO);
       xlabel(handles.axes1,'Weg[mm]','FontSize',AXES_ZIHAO);ylabel(handles.axes1,'Kraft[kN]','FontSize',AXES_ZIHAO);
       title(handles.axes1,get(handles.edit1,'String'),'FontSize',AXES_ZIHAO);
       close(t1);
        msgbox('ͼƬ������ϣ��뱣��ͼ��');
        set(handles.pushbutton2,'Enable','on');
% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
global  newpath MP_FINAL POINT1_INDEX Y1_INDEX  X0 X1 Y1 zihao
handles=guihandles;
guidata(hObject,handles);


[filename1,pathname1]=uiputfile({'*.jpg','JPG files';'*.bmp','BMP files'},'����',newpath);%�������
if isequal(filename1,0)||isequal(pathname1,0)
    return
else
    
    h=figure;
 set(h,'visible','off');
  
    set(h,'position',[100,100,1300,800]); 
     plot(MP_FINAL(:,2),MP_FINAL(:,3)./1000,'linewidth',2);
hold on
plot([X0,X1],[0,Y1]/1000,'--','linewidth',2,'Color','r'); %��һ��������
 %plot(handles.axes1,X_CRUVE,Y_CRUVE/1000)
plot(MP_FINAL(POINT1_INDEX,2),MP_FINAL(POINT1_INDEX,3)/1000, 'o', 'markerfacecolor', [ 1, 0, 0 ]) %��һ�����ݵ�
plot([0,MP_FINAL(POINT1_INDEX,2)],[MP_FINAL(POINT1_INDEX,3)/1000,MP_FINAL(POINT1_INDEX,3)/1000],'--','linewidth',2,'Color','r')%��һ��������
plot(MP_FINAL(Y1_INDEX,2),MP_FINAL(Y1_INDEX,3)/1000, 'o', 'markerfacecolor', [ 1, 0, 0 ]) %����
plot([0,MP_FINAL(Y1_INDEX,2)],[MP_FINAL(Y1_INDEX,3)/1000,MP_FINAL(Y1_INDEX,3)/1000],'--','linewidth',2,'Color','r') %�ڶ���������
 text(MP_FINAL(POINT1_INDEX,2)+1,MP_FINAL(POINT1_INDEX,3)/1000,[num2str(MP_FINAL(POINT1_INDEX,3),'%.0f'),'N'],'FontSize',zihao); %text��һ�������ֵ
        text(MP_FINAL(Y1_INDEX,2),MP_FINAL(Y1_INDEX,3)/1000+2,[num2str(MP_FINAL(Y1_INDEX,3),'%.0f'),'N'],'FontSize',zihao); %text������ֵ
     axis([0 max(MP_FINAL(:,2))*1.1 0 max(MP_FINAL(:,3))/1000*1.1]);grid on;
    grid minor;
      set(gca,'FontSize',zihao);
       
       xlabel('Weg[mm]','FontSize',zihao);ylabel('Kraft[kN]','FontSize',zihao);
       title(get(handles.edit1,'String'),'FontSize',zihao);
     set(gca,'LineWid',2)
end
sa=strcat(pathname1,filename1);
saveas(h,sa);
close(h);
%set(handles.pushbutton2,'Enable','off');




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


% --- Executes on slider movement.
function slider1_Callback(hObject, eventdata, handles)
global  MP_FINAL POINT1_INDEX Y1_INDEX AXES_ZIHAO X0 X1 Y1 Y_CRUVE
sliderValue = get(handles.slider1,'Value');  
ang = int32(sliderValue); 
for i=1:length(Y_CRUVE)
POINT1_INDEX=find((Y_CRUVE(:,1)-MP_FINAL(1:Y1_INDEX,3))>ang,1);
end

 cla(handles.axes1);     
           plot(handles.axes1,MP_FINAL(:,2),MP_FINAL(:,3)./1000,'linewidth',2);
hold on
plot(handles.axes1,[X0,X1],[0,Y1]/1000,'--','linewidth',2,'Color','r'); %��һ��������
 %plot(handles.axes1,X_CRUVE,Y_CRUVE/1000)
plot(handles.axes1,MP_FINAL(POINT1_INDEX,2),MP_FINAL(POINT1_INDEX,3)/1000, 'o', 'markerfacecolor', [ 1, 0, 0 ])
plot(handles.axes1,[0,MP_FINAL(POINT1_INDEX,2)],[MP_FINAL(POINT1_INDEX,3)/1000,MP_FINAL(POINT1_INDEX,3)/1000],'--','linewidth',2,'Color','r')
plot(handles.axes1,MP_FINAL(Y1_INDEX,2),MP_FINAL(Y1_INDEX,3)/1000, 'o', 'markerfacecolor', [ 1, 0, 0 ])
plot(handles.axes1,[0,MP_FINAL(Y1_INDEX,2)],[MP_FINAL(Y1_INDEX,3)/1000,MP_FINAL(Y1_INDEX,3)/1000],'--','linewidth',2,'Color','r')


    axis(handles.axes1,[0 max(MP_FINAL(:,2))*1.1 0 max(MP_FINAL(:,3))/1000*1.1]);grid on
        text(handles.axes1,MP_FINAL(POINT1_INDEX,2)+1,MP_FINAL(POINT1_INDEX,3)/1000,[num2str(MP_FINAL(POINT1_INDEX,3),'%.0f'),'N'],'FontSize',AXES_ZIHAO);
        text(handles.axes1,MP_FINAL(Y1_INDEX,2),MP_FINAL(Y1_INDEX,3)/1000+2,[num2str(MP_FINAL(Y1_INDEX,3),'%.0f'),'N'],'FontSize',AXES_ZIHAO);
       xlabel(handles.axes1,'Weg[mm]','FontSize',AXES_ZIHAO);ylabel(handles.axes1,'Kraft[kN]','FontSize',AXES_ZIHAO);
       title(handles.axes1,get(handles.edit1,'String'),'FontSize',AXES_ZIHAO);


% --- Executes during object creation, after setting all properties.
function slider1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to slider1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: slider controls usually have a light gray background.
if isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor',[.9 .9 .9]);
end
