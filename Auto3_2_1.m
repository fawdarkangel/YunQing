function varargout = Auto3_2_1(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto3_2_1_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto3_2_1_OutputFcn, ...
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


% --- Executes just before Auto3_2_1 is made visible.
function Auto3_2_1_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
movegui(gcf,'center')

b=load([cd,'\interface\Fahrzeugcode.mat'])
for i=1:length(b.Fahrzeugcode)
Fahrzeugcode{i,1}=b.Fahrzeugcode{i,2};
end
set(handles.Fahrzeugcode,'String',Fahrzeugcode);
handles.output = hObject;

guidata(hObject, handles);




% --- Outputs from this function are returned to the command line.
function varargout = Auto3_2_1_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);

[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','ѡ������','MultiSelect','on');
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('�����ļ�ʧ��');
  return;
elseif length(filename)~=48
    msgbox('������������������Ƿ�Ϊ48������');
    return;
else
    
     t1=waitbar(0,'���ڶ�������');
         for i=1:length(filename)
         Filename{i}=strcat(pathname,filename{i});
         [Type Sheet Format]=xlsfinfo(Filename{i}) ;
         sheet{i}=Sheet;
         MP{i}=xlsread(Filename{i},char(sheet{1,i}(1,4)));
         waitbar(i/length(filename));
          try
       system('taskkill/IM excel.exe');
   end
    end 
     close(t1);
end  
    
      RESOLUTION_HE=500;
  RESOLUTION_WI=1300;
  zihao=20;
    
  n=1;%��������
   if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
 end
  Fileadress=strcat(pathname,'result\');
  t2=waitbar(0,'�������ɱ���ͼƬ');
 COLOR_INDEX=[204 0 0;204 189 0;58 204 0;0 204 180;0 49 204;97 0 204;204 0 151;...
      148 71 56;141 148 56;55 149 66;56 148 139;92 72 132;125 73 131;130 74 74;226 130 226;...
      99 23 99;57 231 78;22 188 42]/255;
  TITLE_NAME={'T��r Vorn Links Zug';'T��r Vorn Links Druck';...
      'T��r Vorn Rechts Zug';'T��r Vorn Rechts Druck';...
      'T��r Hinten Links Zug';'T��r Hinten Links Druck';...
       'T��r Hinten Rechts Zug';'T��r Hinten Rechts Druck';};
  for i=1:(length(filename)/6)
   h(i)=figure;
    set(h(i),'visible','off');
     plot(MP{1,n}(:,1),MP{1,n}(:,2),'linewidth',2,'color',COLOR_INDEX(1,:));
     hold on;
     plot(MP{1,n+1}(:,1),MP{1,n+1}(:,2),'linewidth',2,'color',COLOR_INDEX(2,:));
     plot(MP{1,n+2}(:,1),MP{1,n+2}(:,2),'linewidth',2,'color',COLOR_INDEX(3,:));
     plot(MP{1,n+3}(:,1),MP{1,n+3}(:,2),'linewidth',2,'color',COLOR_INDEX(4,:));
    plot(MP{1,n+4}(:,1),MP{1,n+4}(:,2),'linewidth',2,'color',COLOR_INDEX(5,:));
    plot(MP{1,n+5}(:,1),MP{1,n+5}(:,2),'linewidth',2,'color',COLOR_INDEX(6,:));
          set(h(i),'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]);
   set(h(i),'color','w')
        set(gca,'FontSize',zihao);
        title(TITLE_NAME{i},'FontSize',zihao);
     xlabel('Weg[mm]','FontSize',zihao);ylabel('Kraft[N]','FontSize',zihao);  

     Ym=max([max(MP{1,n}(:,2)) max(MP{1,n+1}(:,2)) max(MP{1,n+2}(:,2)) max(MP{1,n+3}(:,2)) max(MP{1,n+4}(:,2)) ...
       max(MP{1,n+5}(:,2)) ])*1.1;
    Xm=max([max(MP{1,n}(:,1)) max(MP{1,n+1}(:,1)) max(MP{1,n+2}(:,1)) max(MP{1,n+3}(:,1)) max(MP{1,n+4}(:,1)) ...
       max(MP{1,n+5}(:,1)) ])*1.3;
    STAND_X=[0;Xm/1.3];
   STAND_Y=[500;500];
       legend('1','2','3','4', '5','6','Location','SouthEast');
   
   grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 Ym]);
   hold off; 
   sfilename1=[Fileadress,num2str(i),'.jpg'];
     f=getframe(h(i));
           imwrite(f.cdata,sfilename1);
           close(h(i));
     n=n+6; 
    waitbar(i/(length(filename)/6));
end
close(t2);

t3=waitbar(0,'�������ɱ���');


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%����Word����%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
filespec_user=[Fileadress,'report.doc'];
try 
Word=actxGetRunningServer('Word.Application');
catch 
Word=actxserver('Word.Application'); 
end
Word.Visible =0; % ʹwordΪ�ɼ�����set(Word, 'Visible', 1); 
%===��word�ļ������·����û���򴴽�һ���հ��ĵ���========================
if exist(filespec_user,'file')
Document=Word.Documents.Open(filespec_user);
else
Document=Word.Documents.Add;
Document.SaveAs2(filespec_user);
end
 t3=waitbar(0.1);
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;

%===�ĵ���ҳ�߾�===========================================================
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;

headline='III. Einzelergebnis ������';
Content.Start=0; % ��ʼ��Ϊ0������ʾÿ��д�븲��֮ǰ����
Content.Text=headline;
Content.Font.Size=10; % �����С
Content.Font.NameAscii='Arial';

Selection.Start = Content.end; 
Selection.TypeParagraph;% ����һ���µĿն���
Selection.Start = Selection.end; 
Selection.TypeParagraph;% ����һ���µĿն���

He=180*1.0771653543307086614173228346457;
Wi=240*1.9;
biaotihao=10;


Tab1 = Document.Tables.Add(Selection.Range,10,12);
DTI = Document.Tables.Item(1); % �����
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % �����ʵ��
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % ���е��ڿ�����

lc=28.381133333333333333333333333333; %���׻���
column_width = [2.74*lc,1.5*lc,1.76*lc,1.09*lc,1.09*lc,1.09*lc,1.09*lc,1.09*lc,2*lc,1.54*lc,1.75*lc,1.75*lc];

for i = 1:12
DTI.Columns.Item(i).Width = column_width(i);
end
 DTI.Range.Paragraphs.Alignment='wdAlignParagraphCenter';
 DTI.Range.Font.NameAscii='Arial';
 DTI.Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';

 for i=1:3
     DTI.Cell(1,i).Merge(DTI.Cell(2,i)); 
 end
 for i=9:12
     DTI.Cell(1,i).Merge(DTI.Cell(2,i)); 
 end
 m=3;
 for i=1:4
     DTI.Cell(m,1).Merge(DTI.Cell(m+1,1));
       DTI.Cell(m,2).Range.Text = 'Zug';
        DTI.Cell(m+1,2).Range.Text = 'Druck';
     m=m+2;
 end
 
  DTI.Cell(3,12).Merge(DTI.Cell(10,12)); 
    DTI.Cell(3,3).Merge(DTI.Cell(10,3));
  DTI.Cell(1,4).Merge(DTI.Cell(1,8)); 
  m=2;
  for i=4:8
    DTI.Cell(2,i).Range.Text =num2str(m);
    m=m+1;
  end
   DTI.Cell(3,3).Range.Text = '50N bzw.10Nm';
  DTI.Cell(1,1).Range.Text = 'Pr��flinge    �����';
  DTI.Cell(1,2).Range.Text = 'Richtung   ����';
  DTI.Cell(1,3).Range.Text = 'Belastng �غ�';
  DTI.Cell(1,4).Range.Text = 'Weg λ��[mm]';
  DTI.Cell(1,5).Range.Text = 'Weg- Mittelwert λ��ƽ��ֵ[mm]';
  DTI.Cell(1,6).Range.Text = 'Winkel �Ƕ�[��]';
  DTI.Cell(1,7).Range.Text ='Istwert ����ֵ[Nm/��]';
  DTI.Cell(1,8).Range.Text = 'Sollwert ����ֵ[Nm/��]';
  DTI.Cell(3,1).Range.Text = 'Vorn links   ��ǰ��';
  DTI.Cell(5,1).Range.Text = 'Vorn rechts  ��ǰ��';
  DTI.Cell(7,1).Range.Text = 'Hinten links  �����';
  DTI.Cell(9,1).Range.Text = 'Hinten rechts �Һ���';
 DTI.Cell(3,12).Range.Text = '��3.5'; 
   t3=waitbar(0.2);
  
  
  %%%%%%%%%%%%%��������ֵ%%%%%%%%%%%%%%5
  for i=1:length(filename)
    WEG_MAX(i)=max(MP{1,i}(:,1));
  end
  
  for i=1:5
  WEG_OUTPUT(1,i)=WEG_MAX(i+1); %��ǰ��������
  end
  m=8;
  for i=1:5
  WEG_OUTPUT(2,i)=WEG_MAX(m);%��ǰ��ѹ����
  m=m+1;
  end
   m=14;
  for i=1:5
  WEG_OUTPUT(3,i)=WEG_MAX(m);%��ǰ��������
  m=m+1;
  end
  m=20;
  for i=1:5
  WEG_OUTPUT(4,i)=WEG_MAX(m);%��ǰ��ѹ����
  m=m+1;
  end
    m=26;
  for i=1:5
  WEG_OUTPUT(5,i)=WEG_MAX(m);%�����������
  m=m+1;
  end
    m=32;
  for i=1:5
  WEG_OUTPUT(6,i)=WEG_MAX(m);%�����ѹ����
  m=m+1;
  end
    m=38;
  for i=1:5
  WEG_OUTPUT(7,i)=WEG_MAX(m);%�Һ���������
  m=m+1;
  end
    m=44;
  for i=1:5
  WEG_OUTPUT(8,i)=WEG_MAX(m);%�Һ���ѹ����
  m=m+1;
  end
  
  for i=4:8
      for j=3:10
     DTI.Cell(j,i).Range.Text =num2str (WEG_OUTPUT(j-2,i-3),'%.2f'); %�����д��λ��ֵ
      end
  end
   t3=waitbar(0.5);
  AVERANGE=mean(WEG_OUTPUT,2);%����ƽ��ֵ
  WINKLE=(AVERANGE./200).*360./2./pi;%��Ƕ�
  SOLL_WINKLE=10./WINKLE;
  for i=3:10
      DTI.Cell(i,9).Range.Text =num2str (AVERANGE(i-2),'%.2f');
      DTI.Cell(i,10).Range.Text =num2str (WINKLE(i-2),'%.2f');
       DTI.Cell(i,11).Range.Text =num2str (SOLL_WINKLE(i-2),'%.2f');
       if SOLL_WINKLE(i-2)<3.5
             DTI.Cell(i,11).Range.Font.Colorindex='wdRed';
             DTI.Cell(i,11).Range.Font.Bold=1;
       end
  end
  Selection.Start = Content.end;
Selection.TypeParagraph;

headline='Kennzeichnung ˵��:  Nach Pr��fnorm: Ersten Wert streichen, restliche Werte Mittelwertbildung. ';
Selection.Text=headline;
Selection.Font.Size=10; % �����С
Selection.Font.NameAscii='Arial';
Selection.Start = Selection.end;
Selection.TypeParagraph;
  headline='                     ���ݱ�׼Ҫ�󣬵�һ�εĲ�������������á�';
Selection.Text=headline;
Selection.Font.Size=10; % �����С
Selection.Font.NameAscii='Arial';

  Selection.Start = Selection.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
 t3=waitbar(0.7);
for i=1:length(filename)/6
    sfilename1=[Fileadress,num2str(i),'.jpg'];
handle=Selection.InlineShapes.AddPicture(sfilename1);
delete(sfilename1); 

end

 t3=waitbar(0.9);
 Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % �����ĵ�
Word.Quit; % �ر��ĵ�
%%%%%%%%%%%%�������������Ϣ�������ռ�%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='���ڻ�������Ťת����';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
t3=waitbar(1);
close(t3);
winopen([Fileadress,'report.doc']);
 


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)
% hObject    handle to Fahrzeugcode (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns Fahrzeugcode contents as cell array
%        contents{get(hObject,'Value')} returns selected item from Fahrzeugcode


% --- Executes during object creation, after setting all properties.
function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)
% hObject    handle to Fahrzeugcode (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
