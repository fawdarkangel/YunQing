function varargout = Auto6_1(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto6_1_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto6_1_OutputFcn, ...
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


% --- Executes just before Auto6_1 is made visible.
function Auto6_1_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to Auto6_1 (see VARARGIN)
handles=guihandles;
guidata(hObject,handles);

%[a b c]=xlsread('\\faw-vw\fs\org\PE\T-E-VC-2\07_测量组mearusing group\12-数据处理平台\resource\Fahrzeugcode.xlsx','Tabelle1','B:B');
b=load([cd,'\interface\Fahrzeugcode.mat'])
for i=1:length(b.Fahrzeugcode)
Fahrzeugcode{i,1}=b.Fahrzeugcode{i,2};
end

set(handles.Fahrzeugcode,'String',Fahrzeugcode);
Cover = imread('PHOTO.PNG');
axes(handles.axes1);
imshow(Cover);
axis off
% Choose default command line output for Auto6_1
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto6_1 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto6_1_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


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
NIHE_MIN=str2num(get(handles.edit1,'String'));
NIHE_MAX=str2num(get(handles.edit2,'String'));
if isempty(NIHE_MIN)||isempty(NIHE_MAX)
    msgbox('请输入拟合范围');
    return
end

val1=get(handles.popupmenu1,'value');
val2=get(handles.popupmenu5,'value');
switch val1
    case 1
      Fzug_FG=20;
      Szug_FG=10;
      Fzug_LH=17;
      Fdruck_LH=17;
      Szug_LH=15;
      Sdruck_LH=15;
               
    case 2
         Fzug_FG=30;
      Szug_FG=4;
      Fzug_LH=22;
      Fdruck_LH=27;
      Szug_LH=5;
      Sdruck_LH=12;
end
switch val2
    case 1
        [filename,pathname,fileindex]=uigetfile('*.txt;*.dat','选择数据','MultiSelect','on');
    case 2
        [filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据','MultiSelect','on');
end
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;
elseif length(filename)~=24
    msgbox('导入文件失败,缺少某个角度试验数据');
   return;
else
ZIHAO_WENZI=10;%所有文字字号
ZIHAO_TU=20;%所有图片字号
end


 
t1=waitbar(0,'正在读入数据');
switch val2
    case 1         %MTS型号
        for i=1:24
           Filename{i}=strcat(pathname,filename{i});
           fidin=fopen(Filename{i});                               % 打开test2.txt文件
           fidout=fopen('result.txt','w');                       % 创建MKMATLAB.txt文件
           tline=fgetl(fidin); 
           tline=fgetl(fidin); 
           while ~feof(fidin)                                      % 判断是否为文件末尾
                 tline=fgetl(fidin);                                     % 从文件读行
                 if isempty(tline)
                    tline=fgetl(fidin);   
                 end
                 if double(tline(1))>=48&&double(tline(1))<=57       % 判断首字符是否是数值
                    fprintf(fidout,'%s\n\n',tline);                  % 如果是数字行，把此行数据写入文件MKMATLAB.txt
                 continue                                         % 如果是非数字继续下一次循环
                 end
           end
                    fclose(fidout);
                    MK=importdata('result.txt');
                    dat2=MK(:,2);
                    dat3=MK(:,3);
%            [dat1 dat2 dat3]=textread(Filename{i},'%f%f%f','headerlines',5);
           MAX_WEG=find(abs(dat2)==max(abs(dat2)));          
            if dat2(round(length(dat2)/2))<0
                MP{1,i}(:,1)=dat2(1:MAX_WEG)*(-1);
                MP{1,i}(:,2)=dat3(1:MAX_WEG)*(-1);
            else
                MP{1,i}(:,1)=dat2(1:MAX_WEG);
                MP{1,i}(:,2)=dat3(1:MAX_WEG);
            end  
            waitbar(i/24);
        end
    case 2        %Zwick
       for i=1:24
           Filename{i}=strcat(pathname,filename{i});
           [Type Sheet Format]=xlsfinfo(Filename{i}) ;
           sheet{i}=Sheet;
           MP{i}=xlsread(Filename{i},char(sheet{1,i}(1,4)));
           waitbar(i/24);
           try
               system('taskkill/IM excel.exe');
           end
       end       
end
try 
    fclose('all')
    delete('result.txt')
end
 close(t1);
 
if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
end
 t2=waitbar(0,'正在生成报告图片');
 F=1;

for i=1:length(MP)
    for j=1:length(MP{1,i})
    if MP{1,i}(j,2)>=20
        a1(i)=j;      %a1为新生成数据起始数
        break;
    end
    end
    MPfinal{1,i}=MP{1,i}(a1(i):length(MP{1,i}),1:2);
end
 
 %% 线性回归并求斜率K
 for i=1:length(MPfinal)
    for j=1:length(MPfinal{1,i})
MPmin=NIHE_MIN;if MPfinal{1,i}(j,2)>=MPmin
a2(i)=MPfinal{1,i}(j,2);Lmin(i)=j;   %a2为线性回归数据起始点
break;
end
    end
 end
 
 for i=1:length(MPfinal)
    for j=1:length(MPfinal{1,i})
MPmax=NIHE_MAX;if MPfinal{1,i}(j,2)>=MPmax
a3(i)=MPfinal{1,i}(j,2);Lmax(i)=j;   %a2为线性回归数据起始点
break;
end
    end
 end
 
  for i=1:length(MPfinal)
MPx{1,i}=MPfinal{1,i}(Lmin(i):Lmax(i),1:2); %MPx为线性回归数据
  end
  
  for i=1:length(MPfinal)
  [p_1,p_2]=polyfit(MPx{1,i}(:,1),MPx{1,i}(:,2),1);%p1(1,1)为斜率
  p1(i)=p_1(1,1);
MPfinal{1,i}(:,1)=MPfinal{1,i}(:,1)-MPfinal{1,i}(1,1);
  end
   waitbar(1/length(MPfinal));
  for i=1:24
  y{i}=p1(i)*(MPfinal{1,i}(:,1)-1); %第一条斜线
  b2(i)=find((MPfinal{1,i}(:,2))==max(MPfinal{1,i}(:,2)));
  b(i)=-(max(MPfinal{1,i}(:,2))/p1(i)-MPfinal{1,i}(b2(i),1)); %%b为最大值对应横坐标
  
  y2{i}=p1(i)*(MPfinal{1,i}(:,1)-b(i));%第二条斜线
  end
  
 
  %%%%%%%%%%%%%%%%%曲线交点%%%%%%%%%%%%%%%5
  for i=1:length(MPfinal)
%    k1{i}=find(abs(y{i}(:,1)-MPfinal{1,i}(:,2))<1);
% if isempty(k1{i})
%     k1{i}=find(abs(y{i}(:,1)-MPfinal{1,i}(:,2))<10);
% end
% if isempty(k1{i})
%     k1{i}=find(abs(y{i}(:,1)-MPfinal{1,i}(:,2))<30);
% end
% if isempty(k1{i})
%     k1{i}=find(abs(y{i}(:,1)-MPfinal{1,i}(:,2))<50);
% end
% if isempty(k1{i})
%     k1{i}=find(abs(y{i}(:,1)-MPfinal{1,i}(:,2))<100);
% end
k1{i}=find(y{i}(:,1)-MPfinal{1,i}(:,2)>=0,1);
  end
  
   for i=1:length(MPfinal)
       zuobiaox1{i}=MPfinal{1,i}(k1{i}(1,1),1);
  F1_mm{i}=MPfinal{1,i}(k1{i}(1,1),2);
   end
   
   TITLE_NAME={'Lenkhebelarm Zug L1';'Lenkhebelarm Zug L2';'Lenkhebelarm Zug L3';'Lenkhebelarm Zug L4';...
       'Lenkhebelarm Zug R1';'Lenkhebelarm Zug R2';'Lenkhebelarm Zug R3';'Lenkhebelarm Zug R4';...
      'Lenkhebelarm Druck L5';'Lenkhebelarm Druck L6';'Lenkhebelarm Druck L7';'Lenkhebelarm Druck L8';...
      'Lenkhebelarm Druck R5';'Lenkhebelarm Druck R6';'Lenkhebelarm Druck R7';'Lenkhebelarm Druck R8';...
     'Fuehrungsgelenkaufnahme Zug L9';'Fuehrungsgelenkaufnahme Zug L10';'Fuehrungsgelenkaufnahme Zug L11';'Fuehrungsgelenkaufnahme Zug L12';...
       'Fuehrungsgelenkaufnahme Zug R9';'Fuehrungsgelenkaufnahme Zug R10';'Fuehrungsgelenkaufnahme Zug R11';'Fuehrungsgelenkaufnahme Zug R12';};
       
   LEGEND_NAME={'L1';'L2';'L3';'L4';'R1';'R2';'R3';'R4';'L5';'L6';'L7';...
       'L8';'R5';'R6';'R7';'R8';'L9';'L10';'L11';'L12';'R9';'R10';'R11';'R12'};
    Fileadress=strcat(pathname,'result\');

   for i=1:length(MPfinal)
       h(i)=figure;
       set(h(i),'visible','off');
       plot(MPfinal{1,i}(:,1),MPfinal{1,i}(:,2)/1000,'linewidth',2);
       hold on
       
        plot(MPfinal{1,i}(:,1),y{i}/1000,'linewidth',2);
        plot(MPfinal{1,i}(:,1),y2{i}/1000,'linewidth',2);
        plot(zuobiaox1{i},F1_mm{i}/1000,'*r');
         plot(b(i),0,'*r');
         set(h(i),'position',[100,100,1300,800]); 
        z=ceil(max(MPfinal{1,i}(:,2))/1000+5);
        axis([0 inf 0 z]);grid on;grid minor;
        set(gca,'FontSize',ZIHAO_TU);
         xlabel('Weg(mm)','FontSize',ZIHAO_TU);ylabel('Kraft(kN)','FontSize',ZIHAO_TU);title(TITLE_NAME{i},'FontSize',ZIHAO_TU);
           text(zuobiaox1{i},F1_mm{i}/1000,['\leftarrow(',num2str(F1_mm{i},'%.f'),'N)'],'FontSize',ZIHAO_TU);
            text(b(i),0,['(',num2str(b(i),'%.2f'),'mm)'],'FontSize',ZIHAO_TU);
           legend({LEGEND_NAME{i},'S1=1mm','S2'},'FontSize',ZIHAO_TU,'Location','southeast');
           sfilename1=[Fileadress,num2str(i),'-',LEGEND_NAME{i},'.jpg'];
        saveas(h(i),sfilename1);
        close(h(i));
        waitbar((i+1)/(length(MPfinal)+1));
   end
   close(t2);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%生成Word报告%%%%%%%%%%%%%%%%%%%%%%%%%%
      t3=waitbar(0,'正在生成Word报告') ;  
         biaotihao=10;
He=180*1.0771653543307086614173228346457;
Wi=240*1.9;
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
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;
headline='III.1 Zugpruefung Lenkhebelarm 转向臂拉力试验';
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=biaotihao; % 字体大小
Content.Font.NameAscii='Arial';
Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落         
 Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落  

%%建立数据表格
Tab1 = Document.Tables.Add(Selection.Range, 10, 8);
DTI = Document.Tables.Item(1); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
% 设置行高，列宽
lc=28.381133333333333333333333333333; %厘米换算
column_width = [2.44*lc,1.05*lc,1.75*lc,1.75*lc,1.5*lc,3.25*lc,3*lc,2.27*lc];

for i = 1:8
DTI.Columns.Item(i).Width = column_width(i);

end
for i=1:10
    for j=1:8
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
        DTI.Cell(i,j).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
    end
end
for i=1:5
DTI.Cell(1,i).Merge(DTI.Cell(2,i));
end
DTI.Cell(1,6).Merge(DTI.Cell(1,7));
DTI.Cell(1,7).Merge(DTI.Cell(2,8));
DTI.Cell(3,6).Merge(DTI.Cell(10,6));
DTI.Cell(3,7).Merge(DTI.Cell(10,7));
DTI.Cell(3,1).Merge(DTI.Cell(6,1));
DTI.Cell(7,1).Merge(DTI.Cell(10,1));

DTI.Cell(1,1).Range.Text = 'Teil Nummer';
DTI.Cell(1,2).Range.Text = 'Nr.';
DTI.Cell(1,3).Range.Text = 'F1mm(N)';
DTI.Cell(1,3).Select;
Selection.Find.Text='1mm';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(1,4).Range.Text = 'SB(mm)';
DTI.Cell(1,4).Select;
Selection.Find.Text='B';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(1,5).Range.Text = 'FB(N)';
DTI.Cell(1,5).Select;
Selection.Find.Text='B';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(1,6).Range.Text = 'Forderung';
DTI.Cell(2,6).Range.Text = 'F1mm≥FZug,LH(kN)';
DTI.Cell(2,6).Select;
Selection.Find.Text='1mm';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,6).Select;
Selection.Find.Text='Zug,LH';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,7).Range.Text = 'SB≥SZug,LH(mm)';
DTI.Cell(2,7).Select;
Selection.Find.Text='B';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,7).Select;
Selection.Find.Text='Zug,LH';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(1,7).Range.Text = 'Bewertung';
waitbar(0.1);
for i=1:8
    DTI.Cell(i+2,2).Range.Text = LEGEND_NAME{i};
    DTI.Cell(i+2,3).Range.Text = num2str(F1_mm{i},'%.f');
    if F1_mm{i}<Fzug_LH
    DTI.Cell(i+2,3).Range.Font.Colorindex='wdRed';
    DTI.Cell(i+2,3).Range.Font.Bold=1;
    end
    DTI.Cell(i+2,4).Range.Text = num2str(b(i),'%.2f');
     if b(i)<Szug_LH
    DTI.Cell(i+2,4).Range.Font.Colorindex='wdRed';
    DTI.Cell(i+2,4).Range.Font.Bold=1;
    end
    DTI.Cell(i+2,5).Range.Text = num2str(max(MPfinal{1,i}(:,2)),'%.f');
end
DTI.Cell(3,6).Range.Text = ['F1mm>=',num2str(Fzug_LH)];
DTI.Cell(3,6).Select;
Selection.Find.Text='1mm';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,7).Range.Text = ['SB>=',num2str(Szug_LH)];
DTI.Cell(3,7).Select;
Selection.Find.Text='B';
Selection.Find.Execute;
Selection.Font.Subscript= true;

Selection.Start = Content.end;
Selection.TypeParagraph;
waitbar(0.2);
headline='Bemerkung:  F1mm:F bei plastischer Verformung 1mm塑性变形量为1mm时的拉力值';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Find.Text='1MM';
Selection.Find.Execute;
Selection.Font.Subscript= true;
Selection.Start = Content.end;
Selection.TypeParagraph;

headline='             SB:Verformung bis Anriss bzw. Bruch 产生裂纹时变形量';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Find.Text='B';
Selection.Find.Execute;
Selection.Font.Subscript= true;
Selection.Start = Content.end;
Selection.TypeParagraph;

headline='             FB:Kraft bis Anriss bzw. Bruch 产生裂纹时拉力值';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Find.Text='B';
Selection.Find.Execute;
Selection.Font.Subscript= true;
Selection.Start = Content.end;
Selection.TypeParagraph;

headline='             FZug,LH:Zulaessige Zugkraft am Spurstangengelenkenzapfen bei s<1mm';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Find.Text='Zug,LH';
Selection.Find.Execute;
Selection.Font.Subscript= true;
Selection.Start = Content.end;
Selection.TypeParagraph;

headline='             在塑性变形小于1mm 时允许施加在球头销上的力值';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start = Selection.end;
Selection.TypeParagraph;

headline='             SZug,LH:Zulaessige Zugverformung ohne Anriss bzw. Bruch 产生裂纹前允许的拉力变形量';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Find.Text='Zug,LH';
Selection.Find.Execute;
Selection.Font.Subscript= true;
Selection.Start = Content.end;
Selection.TypeParagraph;

InlineShapes=Document.InlineShapes;
for i=1:length(MPfinal)
Teil2address{i}=[Fileadress,num2str(i),'-',LEGEND_NAME{i},'.jpg'];
end

for i=1:8
handle=Selection.InlineShapes.AddPicture(Teil2address{1,i});
InlineShapes.Item(i).Height=He;
InlineShapes.Item(i).Width=Wi;
end
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落    
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落 
waitbar(0.3);

%%%%%%%%%%%%%%%%%%转向臂压力结果%%%%%%%%%%%%%%%%%%%%%%%%%%
headline='III.2 Druckpruefung Lenkhebelarm 转向臂压力试验';
Selection.Text=headline;
Selection.Font.Size=biaotihao; % 字体大小
Selection.Font.NameAscii='Arial';
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落         
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落  

%%建立数据表格
Tab2 = Document.Tables.Add(Selection.Range, 10, 8);
DTI = Document.Tables.Item(2); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
% 设置行高，列宽


for i = 1:8
DTI.Columns.Item(i).Width = column_width(i);

end
for i=1:10
    for j=1:8
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
        DTI.Cell(i,j).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
    end
end
for i=1:5
DTI.Cell(1,i).Merge(DTI.Cell(2,i));
end
DTI.Cell(1,6).Merge(DTI.Cell(1,7));
DTI.Cell(1,7).Merge(DTI.Cell(2,8));
DTI.Cell(3,6).Merge(DTI.Cell(10,6));
DTI.Cell(3,7).Merge(DTI.Cell(10,7));
DTI.Cell(3,1).Merge(DTI.Cell(6,1));
DTI.Cell(7,1).Merge(DTI.Cell(10,1));

DTI.Cell(1,1).Range.Text = 'Teil Nummer';
DTI.Cell(1,2).Range.Text = 'Nr.';
DTI.Cell(1,3).Range.Text = 'F1mm(N)';
DTI.Cell(1,3).Select;
Selection.Find.Text='1mm';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(1,4).Range.Text = 'SB(mm)';
DTI.Cell(1,4).Select;
Selection.Find.Text='B';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(1,5).Range.Text = 'FB(N)';
DTI.Cell(1,5).Select;
Selection.Find.Text='B';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(1,6).Range.Text = 'Forderung';
DTI.Cell(2,6).Range.Text = 'F1mm≥FDruck,LH(kN)';
DTI.Cell(2,6).Select;
Selection.Find.Text='1mm';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,6).Select;
Selection.Find.Text='Druck,LH';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,7).Range.Text = 'SB≥SDruck,LH(mm)';
DTI.Cell(2,7).Select;
Selection.Find.Text='B';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,7).Select;
Selection.Find.Text='Druck,LH';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(1,7).Range.Text = 'Bewertung';

waitbar(0.4);
for i=1:8
    DTI.Cell(i+2,2).Range.Text = LEGEND_NAME{i+8};
    DTI.Cell(i+2,3).Range.Text = num2str(F1_mm{i+8},'%.f');
    if F1_mm{i+8}<Fdruck_LH
    DTI.Cell(i+2,3).Range.Font.Colorindex='wdRed';
    DTI.Cell(i+2,3).Range.Font.Bold=1;
    end
    DTI.Cell(i+2,4).Range.Text = num2str(b(i+8),'%.2f');
     if b(i+8)<Sdruck_LH
    DTI.Cell(i+2,4).Range.Font.Colorindex='wdRed';
    DTI.Cell(i+2,4).Range.Font.Bold=1;
    end
    DTI.Cell(i+2,5).Range.Text = num2str(max(MPfinal{1,i+8}(:,2)),'%.f');
end
DTI.Cell(3,6).Range.Text = ['F1mm>=',num2str(Fdruck_LH)];
DTI.Cell(3,6).Select;
Selection.Find.Text='1mm';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,7).Range.Text = ['SB>=',num2str(Sdruck_LH)];
DTI.Cell(3,7).Select;
Selection.Find.Text='B';
Selection.Find.Execute;
Selection.Font.Subscript= true;

Selection.Start = Content.end;
Selection.TypeParagraph;

headline='Bemerkung:  FDruck,LH:Zulaessige Druckkraft am Spurstangengelenkenzapfen bei s<1mm';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Find.Text='Druck,LH';
Selection.Find.Execute;
Selection.Font.Subscript= true;
Selection.Start = Content.end;
Selection.TypeParagraph;

headline='             在塑性变形小于1mm 时允许施加在球头销上的力值';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start = Selection.end;
Selection.TypeParagraph;

waitbar(0.5);
headline='             SDruck,LH:Zulaessige Druckverformung ohne Anriss bzw. Bruch 产生裂纹前允许的压力变形量';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Find.Text='Druck,LH';
Selection.Find.Execute;
Selection.Font.Subscript= true;
Selection.Start = Content.end;
Selection.TypeParagraph;

InlineShapes=Document.InlineShapes;


for i=9:16
handle=Selection.InlineShapes.AddPicture(Teil2address{1,i});
InlineShapes.Item(i).Height=He;
InlineShapes.Item(i).Width=Wi;
end
Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落      
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落 
waitbar(0.6);
%%%%%%%%%%%%%%%%%%转向臂压力结果%%%%%%%%%%%%%%%%%%%%%%%%%%
headline='III.3 Zugpruefung Fuehrungsgelenkaufnahme 控制臂球头销拉力试验';
Selection.Text=headline;
Selection.Font.Size=biaotihao; % 字体大小
Selection.Font.NameAscii='Arial';
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落    
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落 
 

%%建立数据表格
Tab3 = Document.Tables.Add(Selection.Range, 10, 8);
DTI = Document.Tables.Item(3); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
% 设置行高，列宽


for i = 1:8
DTI.Columns.Item(i).Width = column_width(i);

end
for i=1:10
    for j=1:8
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
        DTI.Cell(i,j).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
    end
end
for i=1:5
DTI.Cell(1,i).Merge(DTI.Cell(2,i));
end
DTI.Cell(1,6).Merge(DTI.Cell(1,7));
DTI.Cell(1,7).Merge(DTI.Cell(2,8));
DTI.Cell(3,6).Merge(DTI.Cell(10,6));
DTI.Cell(3,7).Merge(DTI.Cell(10,7));
DTI.Cell(3,1).Merge(DTI.Cell(6,1));
DTI.Cell(7,1).Merge(DTI.Cell(10,1));

DTI.Cell(1,1).Range.Text = 'Teil Nummer';
DTI.Cell(1,2).Range.Text = 'Nr.';
DTI.Cell(1,3).Range.Text = 'F1mm(N)';
DTI.Cell(1,3).Select;
Selection.Find.Text='1mm';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(1,4).Range.Text = 'SB(mm)';
DTI.Cell(1,4).Select;
Selection.Find.Text='B';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(1,5).Range.Text = 'FB(N)';
DTI.Cell(1,5).Select;
Selection.Find.Text='B';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(1,6).Range.Text = 'Forderung';
DTI.Cell(2,6).Range.Text = 'F1mm≥FZug,FG(kN)';
DTI.Cell(2,6).Select;
Selection.Find.Text='1mm';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,6).Select;
Selection.Find.Text='Zug,FG';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,7).Range.Text = 'SB≥SZug,FG(mm)';
DTI.Cell(2,7).Select;
Selection.Find.Text='B';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,7).Select;
Selection.Find.Text='Zug,FG';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(1,7).Range.Text = 'Bewertung';

waitbar(0.7);
for i=1:8
    DTI.Cell(i+2,2).Range.Text = LEGEND_NAME{i+16};
    DTI.Cell(i+2,3).Range.Text = num2str(F1_mm{i+16},'%.f');
    if F1_mm{i+16}<Fzug_FG
    DTI.Cell(i+2,3).Range.Font.Colorindex='wdRed';
    DTI.Cell(i+2,3).Range.Font.Bold=1;
    end
    DTI.Cell(i+2,4).Range.Text = num2str(b(i+16),'%.2f');
     if b(i+16)<Szug_FG
    DTI.Cell(i+2,4).Range.Font.Colorindex='wdRed';
    DTI.Cell(i+2,4).Range.Font.Bold=1;
    end
    DTI.Cell(i+2,5).Range.Text = num2str(max(MPfinal{1,i+16}(:,2)),'%.f');
end
DTI.Cell(3,6).Range.Text = ['F1mm>=',num2str(Fzug_FG)];
DTI.Cell(3,6).Select;
Selection.Find.Text='1mm';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,7).Range.Text = ['SB>=',num2str(Szug_FG)];
DTI.Cell(3,7).Select;
Selection.Find.Text='B';
Selection.Find.Execute;
Selection.Font.Subscript= true;

Selection.Start = Content.end;
Selection.TypeParagraph;

headline='Bemerkung:  FZug,FG:Zulaessige Zugkraft am Fuehrungsgelenkzapfen bei s<1mm';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Find.Text='Zug,FG';
Selection.Find.Execute;
Selection.Font.Subscript= true;
Selection.Start = Content.end;
Selection.TypeParagraph;

waitbar(0.8);
headline='             在塑性变形小于1mm 时允许施加在球头销上的力值';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start = Selection.end;
Selection.TypeParagraph;

headline='             SZug,FG:Zulaessige Zugverformung ohne Anriss bzw. Bruch 产生裂纹前允许的拉力变形量';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Find.Text='Zug,FG';
Selection.Find.Execute;
Selection.Font.Subscript= true;
Selection.Start = Content.end;
Selection.TypeParagraph;

InlineShapes=Document.InlineShapes;


for i=17:24
handle=Selection.InlineShapes.AddPicture(Teil2address{1,i});
InlineShapes.Item(i).Height=He;
InlineShapes.Item(i).Width=Wi;
end
waitbar(0.9);
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档

for i=1:24
    delete(Teil2address{1,i});
end
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='转向节静刚度试验';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
waitbar(1);
close(t3);
winopen([Fileadress,'report.doc']);


% --- Executes on selection change in Fahrzeugcode.

% --- Executes on selection change in fahrzeugcode.
function fahrzeugcode_Callback(hObject, eventdata, handles)

function fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu5.
function popupmenu5_Callback(hObject, eventdata, handles)

function popupmenu5_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



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


% --- Executes on selection change in listbox1.
function listbox1_Callback(hObject, eventdata, handles)
% handles=guihandles;
% guidata(hObject,handles);
global  MPfinal_pre y_pre y2_pre F1_mm_pre b_pre zuobiaox1_pre

ZIHAO_TU=10;
CHOOSE=get(handles.listbox1,'Value');
i=CHOOSE;
 plot(handles.axes1,MPfinal_pre{1,i}(:,1),MPfinal_pre{1,i}(:,2)/1000,'linewidth',2);
       hold on
       
        plot(handles.axes1,MPfinal_pre{1,i}(:,1),y_pre{i}/1000,'linewidth',2);
        plot(handles.axes1,MPfinal_pre{1,i}(:,1),y2_pre{i}/1000,'linewidth',2);
        plot(handles.axes1,zuobiaox1_pre{i},F1_mm_pre{i}/1000,'*r');
         plot(handles.axes1,b_pre(i),0,'*r');        
        z=ceil(max(MPfinal_pre{1,i}(:,2))/1000+5);
        axis(handles.axes1,[0 inf 0 z]);grid on;grid minor;        
         xlabel(handles.axes1,'Weg(mm)','FontSize',ZIHAO_TU);ylabel(handles.axes1,'Kraft(kN)','FontSize',ZIHAO_TU);
           text(handles.axes1,zuobiaox1_pre{i},F1_mm_pre{i}/1000,['\leftarrow(',num2str(F1_mm_pre{i},'%.f'),'N)'],'FontSize',ZIHAO_TU);
            text(handles.axes1,b_pre(i),0,['(',num2str(b_pre(i),'%.2f'),'mm)'],'FontSize',ZIHAO_TU);
          hold off

function listbox1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
clear global MPfinal_pre y_pre y2_pre F1_mm_pre b_pre zuobiaox1_pre
global  MPfinal_pre y_pre y2_pre F1_mm_pre b_pre zuobiaox1_pre
% handles=guihandles;
% guidata(hObject,handles);
NIHE_MIN=str2num(get(handles.edit1,'String'));
NIHE_MAX=str2num(get(handles.edit2,'String'));
if isempty(NIHE_MIN)||isempty(NIHE_MAX)
    msgbox('请输入拟合范围');
    return
end

[filename,pathname,fileindex]=uigetfile('*.txt;*.dat','选择数据','MultiSelect','on');
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;
else
ZIHAO_WENZI=10;%所有文字字号
ZIHAO_TU=20;%所有图片字号
end
t1=waitbar(0,'正在导入数据');
for i=1:length(filename)
           Filename{i}=strcat(pathname,filename{i});
           fidin=fopen(Filename{i});                               % 打开test2.txt文件
           fidout=fopen('result.txt','w');                       % 创建MKMATLAB.txt文件
           tline=fgetl(fidin); 
           tline=fgetl(fidin); 
           while ~feof(fidin)                                      % 判断是否为文件末尾
                 tline=fgetl(fidin);                                     % 从文件读行
                 if isempty(tline)
                    tline=fgetl(fidin);   
                 end
                 if double(tline(1))>=48&&double(tline(1))<=57       % 判断首字符是否是数值
                    fprintf(fidout,'%s\n\n',tline);                  % 如果是数字行，把此行数据写入文件MKMATLAB.txt
                 continue                                         % 如果是非数字继续下一次循环
                 end
           end
                    fclose(fidout);
                    MK=importdata('result.txt');
                    dat2=MK(:,2);
                    dat3=MK(:,3);
%            [dat1 dat2 dat3]=textread(Filename{i},'%f%f%f','headerlines',5);
           MAX_WEG=find(abs(dat2)==max(abs(dat2)));          
            if dat2(round(length(dat2)/2))<0
                MP_pre{1,i}(:,1)=dat2(1:MAX_WEG)*(-1);
                MP_pre{1,i}(:,2)=dat3(1:MAX_WEG)*(-1);
            else
                MP_pre{1,i}(:,1)=dat2(1:MAX_WEG);
                MP_pre{1,i}(:,2)=dat3(1:MAX_WEG);
            end  
            waitbar(i/length(filename));
        end



for i=1:length(MP_pre)
    for j=1:length(MP_pre{1,i})
    if MP_pre{1,i}(j,2)>=20
        a1(i)=j;      %a1为新生成数据起始数
        break;
    end
    end
    MPfinal_pre{1,i}=MP_pre{1,i}(a1(i):length(MP_pre{1,i}),1:2);
end
 %% 线性回归并求斜率K
 for i=1:length(MPfinal_pre)
    for j=1:length(MPfinal_pre{1,i})
MPmin=NIHE_MIN;if MPfinal_pre{1,i}(j,2)>=MPmin
a2(i)=MPfinal_pre{1,i}(j,2);Lmin(i)=j;   %a2为线性回归数据起始点
break;
end
    end
 end
 
 for i=1:length(MPfinal_pre)
    for j=1:length(MPfinal_pre{1,i})
MPmax=NIHE_MAX;if MPfinal_pre{1,i}(j,2)>=MPmax
a3(i)=MPfinal_pre{1,i}(j,2);Lmax(i)=j;   %a2为线性回归数据起始点
break;
end
    end
 end
 
 for i=1:length(MPfinal_pre)
MPx{1,i}=MPfinal_pre{1,i}(Lmin(i):Lmax(i),1:2); %MPx为线性回归数据
  end
  
  for i=1:length(MPfinal_pre)
  [p_1,p_2]=polyfit(MPx{1,i}(:,1),MPx{1,i}(:,2),1);%p1(1,1)为斜率
  p1(i)=p_1(1,1);
MPfinal_pre{1,i}(:,1)=MPfinal_pre{1,i}(:,1)-MPfinal_pre{1,i}(1,1);
  end
  
    for i=1:length(filename)
  y_pre{i}=p1(i)*(MPfinal_pre{1,i}(:,1)-1); %第一条斜线
  b2_pre(i)=find((MPfinal_pre{1,i}(:,2))==max(MPfinal_pre{1,i}(:,2)));
  b_pre(i)=-(max(MPfinal_pre{1,i}(:,2))/p1(i)-MPfinal_pre{1,i}(b2_pre(i),1)); %%b为最大值对应横坐标
  
  y2_pre{i}=p1(i)*(MPfinal_pre{1,i}(:,1)-b_pre(i));%第二条斜线
  end
  
 
  %%%%%%%%%%%%%%%%%曲线交点%%%%%%%%%%%%%%%5
  for i=1:length(MPfinal_pre)
k1{i}=find(y_pre{i}(:,1)-MPfinal_pre{1,i}(:,2)>=0,1);
  end
  
   for i=1:length(MPfinal_pre)
       zuobiaox1_pre{i}=MPfinal_pre{1,i}(k1{i}(1,1),1);
  F1_mm_pre{i}=MPfinal_pre{1,i}(k1{i}(1,1),2);
   end
   try 
    fclose('all')
    delete('result.txt')
   end
   set(handles.listbox1,'Value',1);
set(handles.listbox1,'String',filename);
   close(t1);
