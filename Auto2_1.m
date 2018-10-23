function varargout = Auto2_1(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto2_1_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto2_1_OutputFcn, ...
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


% --- Executes just before Auto2_1 is made visible.
function Auto2_1_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
Cover = imread('Auto2_1.png');
axes(handles.axes1);
imshow(Cover);
axis off
movegui(gcf,'center')

b=load([cd,'\interface\Fahrzeugcode.mat']);
for i=1:length(b.Fahrzeugcode)
Fahrzeugcode{i,1}=b.Fahrzeugcode{i,2};
end
set(handles.Fahrzeugcode,'String',Fahrzeugcode);

% Choose default command line output for Auto2_1
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto2_1 wait for user response (see UIRESUME)
% uiwait(handles.figure1);
% --- Executes during object creation, after setting all properties.
function axes1_CreateFcn(hObject, eventdata, handles)
%imshow(imread('Auto2_1.PNG'));



function varargout = Auto2_1_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;

% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu1 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu1


% --- Executes during object creation, after setting all properties.
function popupmenu1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

function geshi
zihao1=20;
 xlabel('Weg(mm)','FontSize',zihao1);ylabel('Kraft(N)','FontSize',zihao1);  
   legend('Teil 1#','Teil 2#','Teil 3#','Location','NorthWest');

% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
global Fileadress;
handles=guihandles;
guidata(hObject,handles);

 list=get(handles.popupmenu1,'String');
 val1=get(handles.popupmenu1,'Value');

  if val1==1
      msgbox('请选择试验项');
      return
  
  elseif val1==2    
      
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据','MultiSelect','on');

if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;
elseif length(filename)~=72
    msgbox('导入文件失败,缺少某个角度试验数据');
   return;
else
zihao=20;%所有图标的字号
end
 if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
 end
    
t1=waitbar(0,'正在读入数据');
 for i=1:72
         Filename{i}=strcat(pathname,filename{i});
         [Type Sheet Format]=xlsfinfo(Filename{i}) ;
         sheet{i}=Sheet;
         MP{i}=xlsread(Filename{i},char(sheet{1,i}(1,4)));
         waitbar(i/72);
          try
       system('taskkill/IM excel.exe');
   end
 end
         
    
   close(t1) ; 
    t2=waitbar(0,'正在处理，请稍后');
  %% H0 V0 Druck
  RESOLUTION_HE=800;
  RESOLUTION_WI=1300;
 fig_MP1=figure(1);
 for i=1:3
  set(fig_MP1,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
    hold on;
 end
 %set(fig_MP1,'unit','centimeters','position',[0.2,0.2,13.98,7.62]); 
 set(fig_MP1,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP1,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H0° V0° Druck'],'FontSize',zihao);
       Xm=max([max(MP{1,1}(:,1)) max(MP{1,2}(:,1)) max(MP{1,3}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,1}(:,2))*1.1]);
   hold off;   
Fileadress=strcat(pathname,'result\');
    sfilename1=[Fileadress,'1_H0_V0_Druck.jpg'];
     f=getframe(fig_MP1);
           imwrite(f.cdata,sfilename1);
%saveas(fig_MP1,sfilename1);
a=1;
waitbar(a/24);
 a=a+1;  

%% H0 V0 Zug
 fig_MP2=figure(2);
 for i=4:6
  set(fig_MP2,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
 set(fig_MP2,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP2,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H0° V0° Zug'],'FontSize',zihao);
       Xm=max([max(MP{1,4}(:,1)) max(MP{1,5}(:,1)) max(MP{1,6}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,4}(:,2))*1.1]);
   hold off;
    sfilename2=[Fileadress,'2_H0_V0_Zug.jpg'];
      f=getframe(fig_MP2);
           imwrite(f.cdata,sfilename2);
%saveas(fig_MP2,sfilename2);
waitbar(a/24);
 a=a+1;  

%% H0 V5 Druck
 fig_MP3=figure(3);
 for i=7:9
  set(fig_MP3,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
 set(fig_MP3,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP3,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H0° V5° Druck'],'FontSize',zihao);
       Xm=max([max(MP{1,7}(:,1)) max(MP{1,8}(:,1)) max(MP{1,9}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,7}(:,2))*1.1]);
   hold off;
    sfilename3=[Fileadress,'3_H0_V5_Druck.jpg'];
     f=getframe(fig_MP3);
           imwrite(f.cdata,sfilename3);
%saveas(fig_MP3,sfilename3);
waitbar(a/24);
 a=a+1;

%% H0 V5 Zug
 fig_MP4=figure(4);
 for i=10:12
  set(fig_MP4,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
  set(fig_MP4,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP4,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H0° V5° Zug'],'FontSize',zihao);
       Xm=max([max(MP{1,10}(:,1)) max(MP{1,11}(:,1)) max(MP{1,12}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,10}(:,2))*1.1]);
   hold off;
    sfilename4=[Fileadress,'4_H0_V5_Zug.jpg'];
      f=getframe(fig_MP4);
           imwrite(f.cdata,sfilename4);
%saveas(fig_MP4,sfilename4);
waitbar(a/24);
 a=a+1;


%% H0 V-5 Druck
 fig_MP5=figure(5);
 for i=13:15
  set(fig_MP5,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
  set(fig_MP5,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP5,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H0° V-5° Druck'],'FontSize',zihao);
       Xm=max([max(MP{1,13}(:,1)) max(MP{1,14}(:,1)) max(MP{1,15}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,13}(:,2))*1.1]);
   hold off;
    sfilename5=[Fileadress,'5_H0_V-5_Druck.jpg'];
       f=getframe(fig_MP5);
           imwrite(f.cdata,sfilename5);
%saveas(fig_MP5,sfilename5);
waitbar(a/24);
 a=a+1;


%% H0 V-5 Zug
 fig_MP6=figure(6);
 for i=16:18
  set(fig_MP6,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
   set(fig_MP6,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP6,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H0° V-5° Zug'],'FontSize',zihao);
       Xm=max([max(MP{1,16}(:,1)) max(MP{1,17}(:,1)) max(MP{1,18}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,16}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'6_H0_V-5_Zug.jpg'];
     f=getframe(fig_MP6);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP6,sfilename);
waitbar(a/24);
 a=a+1;
 
 
 
%% H30 V0 Druck
 fig_MP7=figure(7);
 for i=19:21
  set(fig_MP7,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
    set(fig_MP7,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP7,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H30° V0° Druck'],'FontSize',zihao);
       Xm=max([max(MP{1,19}(:,1)) max(MP{1,20}(:,1)) max(MP{1,21}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,19}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'7_H30_V0_Druck.jpg'];
      f=getframe(fig_MP7);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP7,sfilename);
waitbar(a/24);
 a=a+1;

%% H30 V0 Zug
 fig_MP8=figure(8);
 for i=22:24
  set(fig_MP8,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
  set(fig_MP8,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP8,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H30° V0° Zug'],'FontSize',zihao);
       Xm=max([max(MP{1,22}(:,1)) max(MP{1,23}(:,1)) max(MP{1,24}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,22}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'8_H30_V0_Zug.jpg'];
     f=getframe(fig_MP8);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP8,sfilename);
waitbar(a/24);
 a=a+1;

%% H30 V5 Druck
 fig_MP9=figure(9);
 for i=25:27
  set(fig_MP9,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
   set(fig_MP9,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP9,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H30° V5° Druck'],'FontSize',zihao);
       Xm=max([max(MP{1,25}(:,1)) max(MP{1,26}(:,1)) max(MP{1,27}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,25}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'9_H30_V5_Druck.jpg'];
     f=getframe(fig_MP9);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP9,sfilename);
waitbar(a/24);
 a=a+1;
 
 
%% H30 V5 Zug
 fig_MP10=figure(10);
 for i=28:30
  set(fig_MP10,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
 set(fig_MP10,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP10,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H30° V5° Zug'],'FontSize',zihao);
       Xm=max([max(MP{1,28}(:,1)) max(MP{1,29}(:,1)) max(MP{1,30}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,28}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'10_H30_V5_Zug.jpg'];
     f=getframe(fig_MP10);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP10,sfilename);
waitbar(a/24);
 a=a+1;
 
 
%% H30 V-5 Druck
 fig_MP11=figure(11);
 for i=31:33
  set(fig_MP11,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
 set(fig_MP11,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP11,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H30° V-5° Druck'],'FontSize',zihao);
       Xm=max([max(MP{1,31}(:,1)) max(MP{1,32}(:,1)) max(MP{1,33}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,31}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'11_H30_V-5_Druck.jpg'];
     f=getframe(fig_MP11);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP11,sfilename);
waitbar(a/24);
 a=a+1;
 
 
 
%% H30 V-5 Zug
 fig_MP12=figure(12);
 for i=34:36
  set(fig_MP12,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
 set(fig_MP12,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP12,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H30° V-5° Zug'],'FontSize',zihao);
       Xm=max([max(MP{1,34}(:,1)) max(MP{1,35}(:,1)) max(MP{1,36}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,34}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'12_H30_V-5_Zug.jpg'];
       f=getframe(fig_MP12);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP12,sfilename);
waitbar(a/24);
 a=a+1;

%% H-30 V0 Druck
 fig_MP13=figure(13);
 for i=37:39
  set(fig_MP13,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
 set(fig_MP13,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP13,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H-30° V0° Druck'],'FontSize',zihao);
       Xm=max([max(MP{1,37}(:,1)) max(MP{1,38}(:,1)) max(MP{1,39}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,37}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'13_H-30_V0_Druck.jpg'];
     f=getframe(fig_MP13);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP13,sfilename);
waitbar(a/24);
 a=a+1;

%% H-30 V0 Zug
 fig_MP14=figure(14);
 for i=40:42
  set(fig_MP14,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
 set(fig_MP14,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP14,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H-30° V0° Zug'],'FontSize',zihao);
       Xm=max([max(MP{1,40}(:,1)) max(MP{1,41}(:,1)) max(MP{1,42}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,40}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'14_H-30_V0_Zug.jpg'];
    f=getframe(fig_MP14);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP14,sfilename);
waitbar(a/24);
 a=a+1;

%% H-30 V5 Druck
 fig_MP15=figure(15);
 for i=43:45
  set(fig_MP15,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
 set(fig_MP15,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP15,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H-30° V5° Druck'],'FontSize',zihao);
       Xm=max([max(MP{1,43}(:,1)) max(MP{1,44}(:,1)) max(MP{1,45}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,43}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'15_H-30_V5_Druck.jpg'];
    f=getframe(fig_MP15);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP15,sfilename);
waitbar(a/24);
 a=a+1;


%% H-30 V5 Zug
 fig_MP16=figure(16);
 for i=46:48
  set(fig_MP16,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
 set(fig_MP16,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP16,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H-30° V5° Zug'],'FontSize',zihao);
       Xm=max([max(MP{1,46}(:,1)) max(MP{1,47}(:,1)) max(MP{1,48}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,46}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'16_H-30_V5_Zug.jpg'];
    f=getframe(fig_MP16);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP16,sfilename);
waitbar(a/24);
 a=a+1;
 
 
%% H-30 V-5 Druck
 fig_MP17=figure(17);
 for i=49:51
  set(fig_MP17,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
 set(fig_MP17,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP17,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H-30° V-5° Druck'],'FontSize',zihao);
       Xm=max([max(MP{1,49}(:,1)) max(MP{1,50}(:,1)) max(MP{1,51}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,49}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'17_H-30_V-5_Druck.jpg'];
     f=getframe(fig_MP17);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP17,sfilename);
waitbar(a/24);
 a=a+1;
 
 
%% H-30 V-5 Zug
 fig_MP18=figure(18);
 for i=52:54
  set(fig_MP18,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
  set(fig_MP18,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP18,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H-30° V-5° Zug'],'FontSize',zihao);
       Xm=max([max(MP{1,52}(:,1)) max(MP{1,53}(:,1)) max(MP{1,54}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,52}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'18_H-30_V-5_Zug.jpg'];
    f=getframe(fig_MP18);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP18,sfilename);
waitbar(a/24);
 a=a+1;

%% H70 V0 Zug
 fig_MP19=figure(19);
 for i=55:57
  set(fig_MP19,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
 set(fig_MP19,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP19,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H70° V0° Zug'],'FontSize',zihao);
       Xm=max([max(MP{1,55}(:,1)) max(MP{1,56}(:,1)) max(MP{1,57}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,55}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'19_H70_V0_Zug.jpg'];
     f=getframe(fig_MP19);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP19,sfilename);
waitbar(a/24);
 a=a+1;

%% H70 V5 Zug
 fig_MP20=figure(20);
 for i=58:60
  set(fig_MP20,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
 set(fig_MP20,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP20,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H70° V5° Zug'],'FontSize',zihao);
       Xm=max([max(MP{1,58}(:,1)) max(MP{1,59}(:,1)) max(MP{1,60}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,58}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'20_H70_V5_Zug.jpg'];
       f=getframe(fig_MP20);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP20,sfilename);
waitbar(a/24);
 a=a+1;


%% H70 V-5 Zug
 fig_MP21=figure(21);
 for i=61:63
  set(fig_MP21,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
 set(fig_MP21,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP21,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H70° V-5° Zug'],'FontSize',zihao);
       Xm=max([max(MP{1,61}(:,1)) max(MP{1,62}(:,1)) max(MP{1,63}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,61}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'21_H70_V-5_Zug.jpg'];
    f=getframe(fig_MP21);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP21,sfilename);
waitbar(a/24);
 a=a+1;
 
 
%% H-70 V0 Zug
 fig_MP22=figure(22);
 for i=64:66
  set(fig_MP22,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
  set(fig_MP22,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP22,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H-70° V0° Zug'],'FontSize',zihao);
       Xm=max([max(MP{1,64}(:,1)) max(MP{1,65}(:,1)) max(MP{1,66}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,64}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'22_H-70_V0_Zug.jpg'];
        f=getframe(fig_MP22);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP22,sfilename);
waitbar(a/24);
 a=a+1;

%% H-70 V5 Zug
 fig_MP23=figure(23);
 for i=67:69
  set(fig_MP23,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
  set(fig_MP23,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP23,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H-70° V5° Zug'],'FontSize',zihao);
       Xm=max([max(MP{1,67}(:,1)) max(MP{1,68}(:,1)) max(MP{1,69}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,67}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'23_H-70_V5_Zug.jpg'];
     f=getframe(fig_MP23);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP23,sfilename);
waitbar(a/24);
 a=a+1;

%% H-70 V-5 Zug
 fig_MP24=figure(24);
 for i=70:72
  set(fig_MP24,'visible','off');
  plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2);
  hold on;
 end
 set(fig_MP24,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]); 
        set(fig_MP24,'color','w')
        set(gca,'FontSize',zihao);
  geshi;
         title(['H-70° V-5° Zug'],'FontSize',zihao);
       Xm=max([max(MP{1,70}(:,1)) max(MP{1,71}(:,1)) max(MP{1,72}(:,1))])*1.1;
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 max(MP{1,70}(:,2))*1.1]);
   hold off;
    sfilename=[Fileadress,'24_H-70_V-5_Zug.jpg'];
       f=getframe(fig_MP24);
           imwrite(f.cdata,sfilename);
%saveas(fig_MP24,sfilename);
waitbar(100);
 
 
 close(t2)
winopen(Fileadress);

set(handles.pushbutton40,'Enable','on');


%% Audi保险杠压力
  elseif val1==3
      run Auto2_1s;
  end





% --------------------------------------------------------------------
function Untitled_1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);


 val1=get(handles.popupmenu1,'Value');
 if val1==1
     msgbox('请选择试验项');
 
 elseif val1==2
dos([cd,'\interface\Auto2_1接口格式.xlsx']);
 elseif val1==3
 dos([cd,'\interface\Auto2_1s接口格式.xlsx']);
 end


% --- Executes on button press in pushbutton40.
function pushbutton40_Callback(hObject, eventdata, handles)
global Fileadress;
handles=guihandles;
guidata(hObject,handles);
filespec_user=[Fileadress,'report.doc'];
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
t4=waitbar(0,'正在生成报告');
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;

%===文档的页边距===========================================================
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;

headline='V.2.Kraft-Weg Diagramme 力-位移曲线';
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=12; % 字体大小

Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落

He=180*1.0771653543307086614173228346457;
Wi=240*1.9;
biaotihao=10;
%% H0 V0 Zug曲线图
REPORT_PHOTO_NAME=[{'2_H0_V0_Zug.jpg'},'1_H0_V0_Druck.jpg','4_H0_V5_Zug.jpg',...
    '3_H0_V5_Druck.jpg','6_H0_V-5_Zug.jpg','5_H0_V-5_Druck.jpg','8_H30_V0_Zug.jpg'...
    ,'7_H30_V0_Druck.jpg','10_H30_V5_Zug.jpg','9_H30_V5_Druck.jpg','12_H30_V-5_Zug.jpg'...
    ,'11_H30_V-5_Druck.jpg','14_H-30_V0_Zug.jpg','13_H-30_V0_Druck.jpg',...
    '16_H-30_V5_Zug.jpg','15_H-30_V5_Druck.jpg','18_H-30_V-5_Zug.jpg',...
    '17_H-30_V-5_Druck.jpg','19_H70_V0_Zug.jpg','20_H70_V5_Zug.jpg',...
    '21_H70_V-5_Zug.jpg','22_H-70_V0_Zug.jpg','23_H-70_V5_Zug.jpg','24_H-70_V-5_Zug.jpg'];
for i=1:24
   REPORT_PHOTO_ADDRESS(i)={[Fileadress,REPORT_PHOTO_NAME{1,i}]};
end
    
 for i=1:24
InlineShapes=Document.InlineShapes;
handle=Selection.InlineShapes.AddPicture(REPORT_PHOTO_ADDRESS{i});
 InlineShapes.Item(i).Height=He;
InlineShapes.Item(i).Width=Wi;
Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落
Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落
waitbar(i/24);
 end

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
%%%%%%%%%%%%输出报告生成信息到公共空间%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='VW拖钩拉力试验';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
close(t4);
set(handles.pushbutton40,'Enable','off');
winopen([Fileadress,'report.doc'])


% --------------------------------------------------------------------
function Untitled_2_Callback(hObject, eventdata, handles)
% hObject    handle to Untitled_2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --------------------------------------------------------------------
function standreport_Callback(hObject, eventdata, handles)
[status,cmdout]=dos('net group "domain admins" /domain'); %判断是否连接公司内网
if status==0
    winopen('\\faw-vw\fs\org\PE\T-E-VC-2\07_测量组mearusing group\12-数据处理平台\标准报告\Auto2_1-VW脱钩拉力.pdf');
elseif status==2
    msgbox('请连接公司内网');
else
    errordlg('错误代码：1001','错误');
end


% --------------------------------------------------------------------
function zhidaoshu_Callback(hObject, eventdata, handles)

function Fahrzeugcode_Callback(hObject, eventdata, handles)

function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
