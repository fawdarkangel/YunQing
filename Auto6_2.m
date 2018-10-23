function varargout = Auto6_2(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto6_2_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto6_2_OutputFcn, ...
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


% --- Executes just before Auto6_2 is made visible.
function Auto6_2_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
movegui(gcf,'center')

b=load([cd,'\interface\Fahrzeugcode.mat']);
for i=1:length(b.Fahrzeugcode)
Fahrzeugcode{i,1}=b.Fahrzeugcode{i,2};
end
set(handles.Fahrzeugcode,'String',Fahrzeugcode);
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto6_2 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto6_2_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)

function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
clear global F_1mm Y_1mm
global F_1mm Y_1mm
NIHE_MIN=str2num(get(handles.edit1,'String'));
NIHE_MAX=str2num(get(handles.edit2,'String')); 
HUOSAIGAN_DIAMETER=str2num(get(handles.edit3,'String')); 
F1MM_VALUE=get(handles.checkbox2,'Value');
if isempty(NIHE_MIN)||isempty(NIHE_MAX)
    msgbox('请输入拟合范围');
    return
end
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据','MultiSelect','on');
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;
else
Teil_number=length(filename);
 t1=waitbar(0,'正在读入数据');
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

%%%%%%%%%%%%%%%数据处理%%%%%%%%%%%%%%%%
RESOLUTION_HE=500;
RESOLUTION_WI=1300;
zihao=20;
if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
end
Fileadress=strcat(pathname,'result\');
t2=waitbar(0,'正在生成报告图片');
 %% 线性回归并求斜率K
 for i=1:length(filename)
    for j=1:length(MP{1,i})
          if MP{1,i}(j,2)>=NIHE_MIN
              a2(i)=MP{1,i}(j,2);Lmin(i)=j;   %a2为线性回归数据起始点
              break;
          end
    end
 end
 
 for i=1:length(filename)
    for j=1:length(MP{1,i})
         if MP{1,i}(j,2)>=NIHE_MAX
            a3(i)=MP{1,i}(j,2);Lmax(i)=j;   %a2为线性回归数据起始点
            break;
          end
    end
 end



for i=1:length(filename)
      MPx{1,i}=MP{1,i}(Lmin(i):Lmax(i),1:2); %MPx为线性回归数据
end
for i=1:length(filename)
     [p_1,p_2]=polyfit(MPx{1,i}(:,1),MPx{1,i}(:,2),1);%p1(1,1)为斜率
     p1(i)=p_1(1,1);  
     B(i)=p_1(1,2);     
     Y_0_5{1,i}=p1(i)*(MP{1,i}(:,1)-0.5);                                                       %0.5mm拟合曲线的Y坐标
     F_0_5_index(i)=find(Y_0_5{1,i}-MP{1,i}(:,2)>=0,1);                               %F0.5mm的Y坐标索引
     F_0_5(i)=MP{1,i}(F_0_5_index(i),2);                                                      %F0.5mm的值
     Yingli0_5(i)=8* F_0_5(i)*280/pi/(HUOSAIGAN_DIAMETER^3);            %0.5mm的应力值 
      Y_025{1,i}=p1(i)*(MP{1,i}(:,1)-0.25);                                                    %0.25mm拟合曲线的Y坐标
      F_025_index(i)=find(Y_025{1,i}-MP{1,i}(:,2)>=0,1);                              %F0.25mm的Y坐标索引      
      F_025(i)=MP{1,i}(F_025_index(i),2);                                                     %0.25mm的力值 
      Yingli0_25(i)=8* F_025(i)*280/pi/(HUOSAIGAN_DIAMETER^3);
end
if get(handles.checkbox2,'Value')==1
    for i=1:length(filename)
          Y_1mm{1,i}=p1(i)*(MP{1,i}(:,1)-1);                                                   %1mm拟合曲线的Y坐标
          F_1mm_index(i)=find(Y_1mm{1,i}-MP{1,i}(:,2)>=0,1);                      %F1mm的Y坐标索引
          F_1mm(i)=MP{1,i}(F_1mm_index(i),2);                                              %F1mm的值     
    end
end
%%%%%%%%%%%%%%%%%%%%%%%%%%出图%%%%%%%%%%%%%%%%%%%%%%%%%%%%
for i=1:length(filename)
    h=figure;
    set(h,'visible','off');
    plot(MP{1,i}(:,1),MP{1,i}(:,2),'linewidth',2)
    hold on;
    plot(MP{1,i}(:,1),Y_025{1,i},'linewidth',1.5);  
    plot(MP{1,i}(:,1),Y_0_5{1,i},'linewidth',1.5);   
    
    if F1MM_VALUE==1
    plot(MP{1,i}(:,1),Y_1mm{1,i},'linewidth',1.5);
    plot(MP{1,i}(F_1mm_index(i),1),MP{1,i}(F_1mm_index(i),2),'o','markerfacecolor',[1 0 0] ,'color','r','MarkerSize',4);     
    text(MP{1,i}(F_1mm_index(i),1),MP{1,i}(F_1mm_index(i),2),['\leftarrow(',num2str(F_1mm(i),'%.f'),'N)'],'FontSize',15);    
    end
    plot(MP{1,i}(F_025_index(i),1),MP{1,i}(F_025_index(i),2),'o','markerfacecolor',[1 0 0] ,'color','r','MarkerSize',4);
    text(MP{1,i}(F_025_index(i),1),MP{1,i}(F_025_index(i),2),['\leftarrow(',num2str(F_025(i),'%.f'),'N)'],'FontSize',15);
    plot(MP{1,i}(F_0_5_index(i),1),MP{1,i}(F_0_5_index(i),2),'o','markerfacecolor',[1 0 0] ,'color','r','MarkerSize',4);
    text(MP{1,i}(F_0_5_index(i),1),MP{1,i}(F_0_5_index(i),2),['\leftarrow(',num2str(F_0_5(i),'%.f'),'N)'],'FontSize',15);
    axis([0 65 0 max(MP{1,i}(:,2))*1.1]);
    grid on;
    set(gca,'FontSize',zihao);
    title(['Teil ',num2str(i),'#'],'FontSize',zihao);
    xlabel('Weg(mm)','FontSize',zihao);ylabel('Kraft(N)','FontSize',zihao);  
    if F1MM_VALUE==1
    legend(['Teil ',num2str(i),'#'],'s1=0.25mm','s2=0.5mm','s3=1mm','Location','SouthEast');
    else
    legend(['Teil ',num2str(i),'#'],'s1=0.25mm','s2=0.5mm','Location','SouthEast');    
    end
    set(h,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]);
    set(h,'color','w');   
    hold off; 
    sfilename1=[Fileadress,num2str(i),'.jpg'];
    f=getframe(h);
    imwrite(f.cdata,sfilename1);
    close(h);    
    waitbar(i/length(filename));
end
close(t2);
%%%%%%%%%%%%%%%%%%%%%%%%%%%报告输出%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
t3=waitbar(0,'正在生成报告');
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
 t3=waitbar(0.1);
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;

%===文档的页边距===========================================================
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;


headline='III. Einzelergebnis 具体结果';
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=10; % 字体大小
Content.Font.NameAscii='Arial';

Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落

He=180*1.0771653543307086614173228346457;
Wi=240*1.9;
biaotihao=10;
lc=28.381133333333333333333333333333; %厘米换算
if F1MM_VALUE==1
Tab1 = Document.Tables.Add(Selection.Range, length(filename)+1,8);
column_width = [1.28*lc,2*lc,2*lc,1.75*lc,2.05*lc,2.95*lc,3.5*lc,2.25*lc];
else 
Tab1 = Document.Tables.Add(Selection.Range, length(filename)+1,7);
column_width = [1.28*lc,2*lc,2*lc,1.75*lc,2.95*lc,3.5*lc,2.25*lc];
end
DTI = Document.Tables.Item(1); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
 t3=waitbar(0.3);
for i = 1:length(column_width)
DTI.Columns.Item(i).Width = column_width(i);
end
DTI.Range.Paragraphs.Alignment='wdAlignParagraphCenter';
DTI.Range.Font.NameAscii='Arial';
DTI.Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(1,1).Range.Text = 'Nr./零件编号';
DTI.Cell(1,2).Range.Text = 'σFilessgrenze/屈服应力[N/mm2]';
DTI.Cell(1,2).Select;
Selection.Find.Text='Filessgrenze';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(1,2).Select;
Selection.Find.Text='2';
Selection.Find.Execute;
Selection.Font.Superscript= true;
DTI.Cell(1,3).Range.Text = 'F0.5mm/0.5mm变形载荷[N]';
DTI.Cell(1,3).Select;
Selection.Find.Text='0.5mm';
Selection.Find.Execute;
Selection.Font.Subscript= true;

DTI.Cell(1,4).Range.Text = 'σ0.5mm/ 0.5mm应力[N/mm2]';
DTI.Cell(1,4).Select;
Selection.Find.Text='0.5mm';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(1,4).Select;
Selection.Find.Text='2';
Selection.Find.Execute;
Selection.Font.Superscript= true;

if F1MM_VALUE==1
DTI.Cell(1,5).Range.Text = 'F1mm/1mm变形载荷[N]';
DTI.Cell(1,5).Select;
Selection.Find.Text='1mm';
Selection.Find.Execute;
Selection.Font.Subscript= true;  
DTI.Cell(1,6).Range.Text = 'Zustand bei Druchbiegung 60mm/60mm变形状态';
DTI.Cell(1,7).Range.Text = 'Forderung/要求';
DTI.Cell(1,8).Range.Text = 'Bewertung/评价';
DTI.Cell(2,7).Merge(DTI.Cell(length(filename)+1,7)); 
DTI.Cell(2,7).Range.Text = 'σFilessgrenze≥500N/mm2;σ0.5mm≥1000N/mm2;Durchbiegung bis zum Bruch≥40mm;F1mm≥';
DTI.Cell(2,7).Select;
Selection.Find.Text='Filessgrenze';
Selection.Find.Execute;
Selection.Font.Subscript= true;  
DTI.Cell(2,7).Select;
Selection.Find.Text='2';
Selection.Find.Execute;
Selection.Font.Superscript= true;
DTI.Cell(2,7).Select;
Selection.Find.Text='2';
Selection.Find.Execute;
Selection.Font.Superscript= true;
DTI.Cell(2,7).Select;
Selection.Find.Text='1mm';
Selection.Find.Execute;
Selection.Font.Subscript= true;  
DTI.Cell(2,7).Select;
Selection.Find.Text='0.5mm';
Selection.Find.Execute;
Selection.Font.Subscript= true;  
else
DTI.Cell(1,5).Range.Text = 'Zustand bei Druchbiegung 60mm/60mm变形状态';
DTI.Cell(1,6).Range.Text = 'Forderung/要求';
DTI.Cell(1,7).Range.Text = 'Bewertung/评价';
DTI.Cell(2,6).Merge(DTI.Cell(length(filename)+1,6)); 
DTI.Cell(2,6).Range.Text = 'σFilessgrenze≥500N/mm2;σ0.5mm≥1000N/mm2;Durchbiegung bis zum Bruch≥40mm';
DTI.Cell(2,6).Select;
Selection.Find.Text='Filessgrenze';
Selection.Find.Execute;
Selection.Font.Subscript= true;  
DTI.Cell(2,6).Select;
Selection.Find.Text='2';
Selection.Find.Execute;
Selection.Font.Superscript= true;
DTI.Cell(2,6).Select;
Selection.Find.Text='2';
Selection.Find.Execute;
Selection.Font.Superscript= true;
DTI.Cell(2,6).Select;
Selection.Find.Text='0.5mm';
Selection.Find.Execute;
Selection.Font.Subscript= true;  
end
 t3=waitbar(0.7);
for i=1:length(filename)
DTI.Cell(i+1,1).Range.Text =[num2str(i),'#'];
DTI.Cell(i+1,2).Range.Text =num2str(Yingli0_25(i),'%.f');
DTI.Cell(i+1,3).Range.Text =num2str(F_0_5(i),'%.f');
DTI.Cell(i+1,4).Range.Text =num2str(Yingli0_5(i),'%.f');
       if Yingli0_5(i)<1000
             DTI.Cell(i+1,4).Range.Font.Colorindex='wdRed';
             DTI.Cell(i+1,4).Range.Font.Bold=1;
       end
end

if F1MM_VALUE==1
  for i=1:length(filename)
      DTI.Cell(i+1,5).Range.Text =num2str(F_1mm(i),'%.f');
      if MP{1,i}(end,2)>1000
         DTI.Cell(i+1,6).Range.Text ='kein Bruch';  
      else
          D=diff(MP{1,i}(:,2));
          INDEX=find(D==min(D));
          DTI.Cell(i+1,6).Range.Text =[num2str(MP{1,i}(INDEX,1),'%.2f'),'mm Bruch'];  
      end
   end
else
   for i=1:length(filename)
      if MP{1,i}(end,2)>1000
         DTI.Cell(i+1,5).Range.Text ='kein Bruch';    
      else
          D=diff(MP{1,i}(:,2));
          INDEX=find(D==min(D));
         DTI.Cell(i+1,5).Range.Text =[num2str(MP{1,i}(INDEX,1),'%.2f'),'mm Bruch'];           
      end
   end
end
t3=waitbar(0.8);
Selection.Start = Content.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
InlineShapes=Document.InlineShapes;
for i=1:length(filename)
    sfilename1=[Fileadress,num2str(i),'.jpg'];
handle=Selection.InlineShapes.AddPicture(sfilename1);
delete(sfilename1); 
end
 t3=waitbar(0.9);

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
%%%%%%%%%%%%输出报告生成信息到公共空间%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='活塞杆弯曲强度试验';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
t3=waitbar(1);
close(t3);
winopen([Fileadress,'report.doc']);
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


% --- Executes on button press in checkbox2.
function checkbox2_Callback(hObject, eventdata, handles)



function edit3_Callback(hObject, eventdata, handles)


function edit3_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
