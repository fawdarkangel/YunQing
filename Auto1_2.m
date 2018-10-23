function varargout = Auto1_2(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto1_2_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto1_2_OutputFcn, ...
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


% --- Executes just before Auto1_2 is made visible.
function Auto1_2_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
movegui(gcf,'center')

b=load([cd,'\interface\Fahrzeugcode.mat'])
for i=1:length(b.Fahrzeugcode)
Fahrzeugcode{i,1}=b.Fahrzeugcode{i,2};
end
set(handles.Fahrzeugcode,'String',Fahrzeugcode);
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto1_2 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto1_2_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);

[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据','MultiSelect','on');

if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;

else
ZIHAO_WENZI=10;%所有文字字号
ZIHAO_TU=16;%所有图片字号
end

 if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
 end
 Fileadress=strcat(pathname,'result\');
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
 t2=waitbar(0,'正在生成报告图片');
 %%%%%%%%%%%%%%%%%%%%%%%%%生成图片%%%%%%%%%%%%%%%%%5
 for i=1:length(filename)
 START_INDEX(i)=find(MP{1,i}(:,2)>0.1,1); %第一个力大于0的值索引下标，即曲线起始点
 MAX_INDEX(i)=find(MP{1,i}(:,2)==max(MP{1,i}(:,2)));
END_INDEX_1=find(MP{1,i}(MAX_INDEX(i):end,2)<0,1)-1;  %所有力小于0的下标

if isempty(END_INDEX_1)
        END_INDEX(i)=length(MP{1,i});  %如果最后力大于0的话，最后一个值为终止点
else
END_INDEX(i)=END_INDEX_1+ MAX_INDEX(i)-1; %最后一个力大于0的下标，即曲线终止点
end
MP_final{1,i}=MP{1,i}(START_INDEX:END_INDEX(i),1:2);                       %截取后的最终曲线


END_INDEX_80(i)=find(MP_final{1,i}(:,2)>80,1)-1;                                %力为80N时下标
MP_final_80{1,i}=MP_final{1,i}(1:END_INDEX_80(i),1:2);                         %力从0到80N数据
END_INDEX_120(i)=find(MP_final{1,i}(:,2)>120,1)-1;                              %力为120N时下标
MP_final_120{1,i}=MP_final{1,i}(END_INDEX_80(i)+1:END_INDEX_120(i),1:2);%力从80N到120N数据

END_INDEX_200(i)=find(MP_final{1,i}(:,2)==max(MP_final{1,i}(:,2)));               %力为最大值时下标
MP_final_200{1,i}=MP_final{1,i}(END_INDEX_120(i)+1:END_INDEX_200(i),1:2);%力为120N至最大值数据
 end
  
   RESOLUTION_HE=600;                                                               %生成图片高度像素
  RESOLUTION_WI=1300;                                                               %生成图片宽度像素
  
 for i=1:length(filename)
     h(i)=figure;
     set(h(i),'visible','off');
 Y120=[120 120];
 Y80=[80 80];
 Xm=max(MP_final{1,i}(:,1));                                                    %求变形最大值
 Ym_INDEX=find(MP_final{1,i}(:,2)==max(MP_final{1,i}(:,2)));    %求力最大值
 X0=[0 Xm];
 x=[0 Xm*1.1 Xm*1.1 0];                                                 %80与120N矩形横坐标
y=[80 80 120 120];                                                          %80与120N矩形纵坐标

plot(MP_final{1,i}(:,1),MP_final{1,i}(:,2),'LineWidth',2);          %画零件Weg-Kraft曲线
hold on

 set(h(i),'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]);
   set(h(i),'color','w')
        set(gca,'FontSize',ZIHAO_TU);
         xlabel('Weg/标称应变[mm]','FontSize',ZIHAO_TU);ylabel('Kraft/标准载荷[N]','FontSize',ZIHAO_TU);  
           axis([0 Xm*1.1 0 240]);
 %%%%%%%%%判断曲线是否有拐点，有则在图上注明%%%%%%%%%%%%
      if all(diff(MP_final_80{1,i}(:,2))>0)
          text(max(MP_final{1,i}(:,1))/2,30,'Kein Wendepunkt unter 80N','FontWeight','bold','FontSize',ZIHAO_TU);
      else
          text(max(MP_final{1,i}(:,1))/2,30,'Es gibt Wendepunkt unter 80N','color','r','FontWeight','bold','FontSize',ZIHAO_TU);
      end
      if all(diff(MP_final_120{1,i}(:,2))>0)
           text(max(MP_final{1,i}(:,1))/10,100,'Keine negative Steigung ','FontWeight','bold','FontSize',ZIHAO_TU);
      else
          text(max(MP_final{1,i}(:,1))/10,100,'Es gibt negative Steigung','color','r','FontWeight','bold','FontSize',ZIHAO_TU);
      end
       if all(diff(MP_final_200{1,i}(:,2))>=0)
           text(max(MP_final{1,i}(:,1))/3,180,{'>120N Wendepunkt';'nur mit Steigung>=0'},'FontWeight','bold','FontSize',ZIHAO_TU);
      else
          text(max(MP_final{1,i}(:,1))/3,180,{'>120N Es gibt Wendepunkt';'mit Steigung<0'},'color','r','FontWeight','bold','FontSize',ZIHAO_TU);
       end
         grid on; set(gca, 'GridLineStyle' ,'-');
               title(['Kraft/Weg Kurve am MP',num2str(i),' Dachpoliersteifigkeit'],'FontSize',ZIHAO_TU);
               
       B=(MP_final{1,i}(Ym_INDEX,2))-MP_final{1,i}(Ym_INDEX,1)*200/15;                      %B刚度标准15N/mm截距
       X_200N_STAND=[Xm/1.1 Xm*1.05];                                                                     %15N/mm辅助线横坐标
       Y_200N_STAND=200/15.*X_200N_STAND+B;                                                     %15N/mm辅助线纵坐标
               plot(X_200N_STAND,Y_200N_STAND,'--m','LineWidth',2)                              %画15N/mm辅助线
        %K=(MP_final{1,i}(Ym_INDEX,2)-MP_final{1,i}(Ym_INDEX-1,2))/ (MP_final{1,i}(Ym_INDEX,1)-MP_final{1,i}(Ym_INDEX-1,1));
        K=200/Xm;                                                                                                                 %200N实际斜率
       Y_200N_REAL=[190 210];                                                               %200N辅助线横坐标
          X_200N_REAL=[ Y_200N_REAL(1)/K  Y_200N_REAL(2)/K];                 %200N辅助线纵坐标
       plot(X_200N_REAL,Y_200N_REAL,'--g','LineWidth',2);                                                %画拟合刚度曲线
       patch(x,y,'b','linestyle','none','facealpha','0.3');                                                            %画80-120N透明矩形
       legend('Teil','C_1_5_N_/_m_m','C_T_e_i_l','Location','SouthEast');
       hold off;
       
    %STEIFIGKEIT(i)=MP_final{1,i}(Ym_INDEX,2)/MP_final{1,i}(Ym_INDEX,1);
    MAX_VERFORMUNG(i)=Xm;  %求最大变性量
    K200(i)=K;                                   %求刚度值
    BLE_VERFORMUNG(i)=MP_final{1,i}(length(MP_final{1,i}),1);                                            %求残余变形量
        sfilename1=[Fileadress,num2str(i),'.jpg'];
     f=getframe(h(i));
           imwrite(f.cdata,sfilename1);
           close(h(i));
          waitbar(i/length(filename));
 end
       
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%生成Word报告%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
close(t2);
t3=waitbar(0,'正在生成Word报告');
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
t3=waitbar(0.1);
%===文档的页边距===========================================================
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;
biaotihao=10;
headline=['III. Einzelergebnis 具体结果'];
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=biaotihao; % 字体大小
Content.Font.NameAscii='Arial';

Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落

InlineShapes=Document.InlineShapes;
He=180*1.0771653543307086614173228346457;
Wi=240*1.9;
biaotihao=10;

Tab1 = Document.Tables.Add(Selection.Range, length(filename)+1,5);
DTI = Document.Tables.Item(1); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条

lc=28.381133333333333333333333333333; %厘米换算
column_width = [2.24*lc,3.51*lc,2.75*lc,3*lc,3.25*lc];
t3=waitbar(0.3);
for i = 1:5
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:length(filename)+1
    for j=1:5
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
        DTI.Cell(i,j).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
    end
end
t3=waitbar(0.6);
DTI.Cell(1,1).Range.Text = 'Messpunkt';
DTI.Cell(1,2).Range.Text = 'Steifigkeit bei 200N[N/mm]';
DTI.Cell(1,3).Range.Text = 'Soll- Steifigkeit [N/mm]';
DTI.Cell(1,4).Range.Text = 'max.Verformung[mm]';
DTI.Cell(1,5).Range.Text = 'bleibende Verformung[mm]';

DTI.Cell(2,3).Merge(DTI.Cell(length(filename)+1,3)); 
DTI.Cell(2,3).Range.Text = '>=15';



for i=1:length(filename)
DTI.Cell(i+1,1).Range.Text =['MP',num2str(i)];
DTI.Cell(i+1,2).Range.Text =num2str(K200(i),'%.2f');                                                  %输出200N时刚度值
 if K200(i)<15                                                                                                            %判断刚度是否小于15，小于则标红加粗
             DTI.Cell(i+1,2).Range.Font.Colorindex='wdRed';
             DTI.Cell(i+1,2).Range.Font.Bold=1;
       end
DTI.Cell(i+1,4).Range.Text =num2str(MAX_VERFORMUNG(i),'%.2f');                           %输出最大变形
DTI.Cell(i+1,5).Range.Text =num2str(BLE_VERFORMUNG(i),'%.2f');                              %输出残余变形
end

Selection.Start = Content.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
InlineShapes=Document.InlineShapes;
t3=waitbar(0.7);
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
TEST_NAME='Poliersteifigkeit';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
t3=waitbar(1);
close(t3);
winopen([Fileadress,'report.doc']);


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)

function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
