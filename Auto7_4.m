function varargout = Auto7_4(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto7_4_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto7_4_OutputFcn, ...
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


% --- Executes just before Auto7_4 is made visible.
function Auto7_4_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
movegui(gcf,'center')
load([cd,'\interface\Config\Auto7_4_Version.mat'])            %��ȡ�����ļ�

DATA_TYPE=CONFIG_7_4.DATA_TYPE;
DATA_TYPE_KRAFT=CONFIG_7_4.DATA_TYPE_KRAFT;      %��ȡ���ݵڼ���Ϊ��
DATA_TYPE_WEG=CONFIG_7_4.DATA_TYPE_WEG;          %��ȡ���ݵڼ���Ϊλ��
X_LABLE=CONFIG_7_4.X_LABLE;                                 %��ȡ������
Y_LABLE=CONFIG_7_4.Y_LABLE;                                 %��ȡ������


setappdata(0,'CONFIG_7_4',CONFIG_7_4);                             %������д���ڴ�
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto7_4 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto7_4_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on selection change in listbox1.
function listbox1_Callback(hObject, eventdata, handles)
cla(handles.axes1);

MP=getappdata(0,'MP');
CONFIG_7_4=getappdata(0,'CONFIG_7_4');                             %��ȡ�ڴ�����
DATA_TYPE=CONFIG_7_4.DATA_TYPE;
DATA_TYPE_KRAFT=CONFIG_7_4.DATA_TYPE_KRAFT;      %��ȡ���ݵڼ���Ϊ��
DATA_TYPE_WEG=CONFIG_7_4.DATA_TYPE_WEG;          %��ȡ���ݵڼ���Ϊλ��
X_LABLE=CONFIG_7_4.X_LABLE;                                 %��ȡ������
Y_LABLE=CONFIG_7_4.Y_LABLE;                                 %��ȡ������


filename=getappdata(0,'filename');
EveryFigure_CruveNum=getappdata(0,'EveryFigure_CruveNum');

CHOOSE=get(handles.listbox1,'Value');                %listbox��ֵ
i=CHOOSE;

Cruve_Start_Num=1+(i-1)*length(filename)/EveryFigure_CruveNum;


for j=1:EveryFigure_CruveNum    
    plot(handles.axes1,MP{Cruve_Start_Num}(:,1),MP{Cruve_Start_Num}(:,2));
 hold on;
    Cruve_Start_Num=Cruve_Start_Num+1;
end
hold off;
grid on;
datacursormode on ;

xlabel(handles.axes1,X_LABLE,'FontSize',20);
ylabel(handles.axes1,Y_LABLE,'FontSize',20);


% --- Executes during object creation, after setting all properties.
function listbox1_CreateFcn(hObject, eventdata, handles)


if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit1_Callback(hObject, eventdata, handles)

function edit1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)

      CONFIG_7_4=getappdata(0,'CONFIG_7_4');                             %��ȡ�ڴ�����
      DATA_TYPE=CONFIG_7_4.DATA_TYPE;
      DATA_TYPE_KRAFT=CONFIG_7_4.DATA_TYPE_KRAFT;      %��ȡ���ݵڼ���Ϊ��
      DATA_TYPE_WEG=CONFIG_7_4.DATA_TYPE_WEG;          %��ȡ���ݵڼ���Ϊλ��
if isempty(get(handles.edit1,'String'))
    errordlg('������ÿ��ͼ��������','����');
    return
else   
    EveryFigure_CruveNum=str2double(get(handles.edit1,'String'));         %��ȡÿ��ͼƬ��������   
     setappdata(0,'EveryFigure_CruveNum',EveryFigure_CruveNum)     
 switch DATA_TYPE
     case 1     
    [filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','ѡ������','MultiSelect','on');
     if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
            msgbox('�����ļ�ʧ��');
            return;
     end    
        if rem(length(filename)/EveryFigure_CruveNum,1)~=0
          errordlg('����������ͼƬ������������˶�ԭʼ��������','����');
          return
     end
         t1=waitbar(0,'���ڶ�������');

     for i=1:length(filename)
        Filename{i}=strcat(pathname,filename{i});
        [Type Sheet Format]=xlsfinfo(Filename{i}) ;
        sheet{i}=Sheet;
        MP_MITTLE{i}=xlsread(Filename{i},char(sheet{1,i}(1,4)));
        MP{i}(:,1)=MP_MITTLE{i}(:,DATA_TYPE_WEG);
        MP{i}(:,2)=MP_MITTLE{i}(:,DATA_TYPE_KRAFT);
        waitbar(i/length(filename));
        try
            system('taskkill/IM excel.exe');
        end
     end
     
     case 2
        [filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','ѡ������','MultiSelect','on');
        if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
            msgbox('�����ļ�ʧ��');
            return;
        end       
           if rem(length(filename)/EveryFigure_CruveNum,1)~=0
          errordlg('����������ͼƬ������������˶�ԭʼ��������','����');
          return
     end
        t1=waitbar(0,'���ڶ�������');
        for i=1:length(filename)
            Filename{i}=strcat(pathname,filename{i});
            fidin=fopen(Filename{i});                               % ��test2.txt�ļ�
            fidout=fopen('result.txt','w');                       % ����MKMATLAB.txt�ļ�
            tline=fgetl(fidin);
            tline=fgetl(fidin);
            while ~feof(fidin)                                      % �ж��Ƿ�Ϊ�ļ�ĩβ
                tline=fgetl(fidin);                                     % ���ļ�����
                if isempty(tline)
                    tline=fgetl(fidin);
                end
                if double(tline(1))>=48&&double(tline(1))<=57       % �ж����ַ��Ƿ�����ֵ
                    fprintf(fidout,'%s\n\n',tline);                  % ����������У��Ѵ�������д���ļ�MKMATLAB.txt
                    continue                                         % ����Ƿ����ּ�����һ��ѭ��
                end
            end
            fclose(fidout);
            MK=importdata('result.txt');
            MP{i}(:,1)=MK(:,DATA_TYPE_WEG);
            MP{i}(:,2)=MK(:,DATA_TYPE_KRAFT);
            try
            delete('result.txt');
            end
            waitbar(i/length(filename));
        end
         
         
         
         
         
         
 end
     
     
    close(t1);
     setappdata(0,'MP',MP);
    setappdata(0,'filename',filename);
    setappdata(0,'pathname',pathname);
    setappdata(0,'Filename',Filename);   
    set(handles.listbox1,'Value',1);
    for i=1:1000
        FigureNumber(i)={['Figure',num2str(i)]};             %��ʼ��LIST����
    end
    List_FigureNumber=FigureNumber(1:length(filename)/EveryFigure_CruveNum);  %ͼƬ�������������List��String
    set(handles.listbox1,'String',List_FigureNumber);
    msgbox('���ݵ���ɹ�');

     
     
end




    


% --------------------------------------------------------------------
function Menu1_Callback(hObject, eventdata, handles)



% --------------------------------------------------------------------
function Menu1_1_Callback(hObject, eventdata, handles)
run Auto7_4_Configuration
