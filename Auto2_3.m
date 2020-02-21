function varargout = Auto2_3(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto2_3_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto2_3_OutputFcn, ...
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


% --- Executes just before Auto2_3 is made visible.
function Auto2_3_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
movegui(gcf,'center')

b=load([cd,'\interface\Fahrzeugcode.mat']);
for i=1:length(b.Fahrzeugcode)
Fahrzeugcode{i,1}=b.Fahrzeugcode{i,2};
end
set(handles.Fahrzeugcode,'String',Fahrzeugcode);

load([cd,'\interface\Config\Auto2_3_Config.mat'])            %读取配置文件
setappdata(0,'AUTO_2_3CONFIG',CONFIG);

handles.output = hObject;

% Update handles structure
guidata(hObject, handles);




% --- Outputs from this function are returned to the command line.
function varargout = Auto2_3_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)


function popupmenu1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)

function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



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


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
if isempty(get(handles.edit1,'String'))
    msgbox('请输入项目名称');
    return
end
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择标题文件');

if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
    return;
else
    Filename=strcat(pathname,filename);
    [Type Sheet Format]=xlsfinfo(Filename) ;
    sheet=Sheet;
    [NUM ROW STAND_TITLE]=xlsread(Filename,char(sheet(1,1)));  
    a=STAND_TITLE(:,1);
    setappdata(0,'STAND_TITLE',a);
    setappdata(0,'SOLL_WERT',NUM);
    msgbox('标题导入成功');
    set(handles.pushbutton2,'Enable','on');
    
end



function popupmenu3_Callback(hObject, eventdata, handles)

function popupmenu3_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu4.
function popupmenu4_Callback(hObject, eventdata, handles)

function popupmenu4_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
DATA_TYPE_KRAFT=get(handles.popupmenu3,'value');      %读取数据第几列为力
DATA_TYPE_WEG=get(handles.popupmenu4,'value');          %读取数据第几列为位移

switch get(handles.popupmenu1,'value')
    case 1                                                                             %Zwick多Sheet
        [filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据');
        if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
            msgbox('导入文件失败');
            return;
        else
            Filename=strcat(pathname,filename);
            [Type Sheet Format]=xlsfinfo(Filename)
            if length(getappdata(0,'STAND_TITLE'))~=length(Sheet)
                msgbox('标题数量与数据数量不符，请检查数据或重新导入标题')
                return
            end
            t1=waitbar(0,'正在读取数据');
            for i=1:length(Sheet)
                MK=xlsread(Filename,Sheet{i});
                MP{i}(:,1)=MK(:,DATA_TYPE_WEG);
                MP{i}(:,2)=MK(:,DATA_TYPE_KRAFT);
                waitbar(i/(length(Sheet)));
            end
        end
        set(handles.listbox1,'String',Sheet(1:end));
    case 2
        [filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据','MultiSelect','on');
        if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
           msgbox('导入文件失败');
        return;
      else
          t1=waitbar(0,'正在读入数据');
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
          set(handles.listbox1,'String',filename);
        end
    case 3                                                                             %高铁
        [filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx;*.txt','选择数据','MultiSelect','on');
        if length(getappdata(0,'STAND_TITLE'))~=length(filename)
            msgbox('标题数量与数据数量不符，请检查数据或重新导入标题')
            return
        end
        if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
            msgbox('导入文件失败');
            return;
        end
        t1=waitbar(0,'正在读入数据');
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
            MP{i}(:,1)=MK(:,DATA_TYPE_WEG);
            MP{i}(:,2)=MK(:,DATA_TYPE_KRAFT);
            try
                delete('result.txt');
            end
            waitbar(i/length(filename));
        end
        set(handles.listbox1,'String',filename);
end
close(t1);
t2=waitbar(0,'正在处理数据');
for i=1:length(MP)
    b=diff(MP{i}(:,2));
    INDEX=find(b<-10);
    if isempty(INDEX)
        IST_INDEX(i)=find(MP{i}(:,2)==max(MP{i}(:,2)),1);
        IST_WERT(i)=max(MP{i}(:,2));
    else
        IST_INDEX(i)=min(INDEX)-2;
        IST_WERT(i)=MP{i}(max(IST_INDEX(i)),2);
    end
    waitbar(i/length(MP));
end
close(t2);
setappdata(0,'Auto2_3_IST_WERT',IST_WERT);
setappdata(0,'Auto2_3_IST_INDEX',IST_INDEX);
setappdata(0,'Auto2_3_MP',MP);
setappdata(0,'Auto2_3_pathname',pathname);
set(handles.pushbutton3,'Enable','on');
set(handles.listbox1,'Value',1);
msgbox('数据导入成功');
    


% --- Executes on selection change in listbox1.
function listbox1_Callback(hObject, eventdata, handles)
cla(handles.axes1);
CONFIG=getappdata(0,'AUTO_2_3CONFIG');
STAND_TITLE=getappdata(0,'STAND_TITLE');
SOLL_WERT=getappdata(0,'SOLL_WERT');
MP=getappdata(0,'Auto2_3_MP');
IST_WERT=getappdata(0,'Auto2_3_IST_WERT');
IST_INDEX=getappdata(0,'Auto2_3_IST_INDEX');

CHOOSE=get(handles.listbox1,'Value');                %listbox的值
i=CHOOSE;

ZIHAO_TU_YULAN=CONFIG.FONTSIZE;
TITLEFONTSIZE=CONFIG.TITLEFONTSIZE;


plot(handles.axes1,MP{i}(:,1),MP{i}(:,2),'linewidth',2);

hold on

plot(handles.axes1,MP{i}(IST_INDEX(i),1),MP{i}(IST_INDEX(i),2),'*','Color','r')
hold off
datacursormode on

set(handles.edit2,'String',num2str(SOLL_WERT(i),'%.1f'));
set(handles.edit3,'String',num2str(IST_WERT(i),'%.1f'));
xlabel(handles.axes1,'Weg/位移[mm]','FontSize',ZIHAO_TU_YULAN)
ylabel(handles.axes1,'Kraft/力[N]','FontSize',ZIHAO_TU_YULAN)
title(handles.axes1,STAND_TITLE{i},'FontSize',TITLEFONTSIZE)
 axis(handles.axes1,[0 max(MP{i}(:,1))*1.05 0 max(MP{i}(:,2))*1.1]);
 
 
 
function listbox1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
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


% --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)
PROJECT=get(handles.edit1,'String');
CONFIG=getappdata(0,'AUTO_2_3CONFIG');
STAND_TITLE=getappdata(0,'STAND_TITLE');
SOLL_WERT=getappdata(0,'SOLL_WERT');
MP=getappdata(0,'Auto2_3_MP');
IST_WERT=getappdata(0,'Auto2_3_IST_WERT');
IST_INDEX=getappdata(0,'Auto2_3_IST_INDEX');
ZIHAO_TU_YULAN=CONFIG.FONTSIZE/2;
TITLEFONTSIZE=CONFIG.TITLEFONTSIZE/2;
t1=waitbar(0,'正在生成图片')
pathname=getappdata(0,'Auto2_3_pathname');
     if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
  end
   filename=strcat(pathname,'result\');%合成保存图片路径Ｌ   
     file_usr=strcat(cd,'\model\Audi保险杠拉拔力.pptx');
 copy_usr=['copy ','"',file_usr,'"'] ;
%filespec_user=strcat(pathname,['result\',PROJECT,'.pptx']);
%19.10.31修改，更改文件名，替换文件名中'/'字符串，防止保存文件时出错
this_filename=strrep(PROJECT,'\','_');
this_filename=strrep(this_filename,'/','_');
filespec_user=strcat(pathname,['result\',this_filename,'.pptx']);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
copy_tal=['"',filespec_user,'"'];
xyz=[copy_usr,' ',copy_tal];
dos(xyz);



zihao=CONFIG.FONTSIZE*2.5;
TITLEFONTSIZE=CONFIG.TITLEFONTSIZE*3.5;

for i=1:length(MP)
    h=figure(i);
    plot(MP{i}(:,1),MP{i}(:,2),'linewidth',2);
    set(h,'visible','off');
    %set(h,'unit','centimeters','position',[0.2,0.2,9.2,14.05]);
    set(h,'position',[100,100,CONFIG.Figure_Width,CONFIG.Figure_Height]);
    set(h,'color','w')
    set(gca,'FontSize',zihao);
    grid on;
    xlabel('Weg/位移[mm]','FontSize',zihao)
    ylabel('Kraft/力[N]','FontSize',zihao)
    title(STAND_TITLE{i},'FontSize',TITLEFONTSIZE)
    axis([0 max(MP{i}(:,1))*1.05 0 max(MP{i}(:,2))*1.1]);
   
       sfilename=[filename,'MP' num2str(i) '.jpg'];
       %saveas(h,sfilename);
           f=getframe(h);
           imwrite(f.cdata,sfilename); 
           waitbar(i/length(MP));
close(h);
end
close(t1)

t2=waitbar(0,'正在生成报告');

try
     Powerpoint = actxGetRunningServer('Powerpoint.Application');
 catch
     Powerpoint = actxserver('Powerpoint.Application'); 
 end
 %Powerpoint.Visible = 0;    
 set(Powerpoint, 'Visible', 1); 
 if exist(filespec_user,'file')
     Presentation = Powerpoint.Presentation.Open(filespec_user);
 else
     Presentation = Powerpoint.Presentation.Add;
         Presentation.SaveAs(filespec_user);
 end
     Slides = Powerpoint.ActivePresentation.Slides; 
 Slides1=Slides.Item(1);
 Slides1.Copy;
 for i=1:length(MP)-1
          Slides.Paste;
 end

 for i=1:length(MP)
     Slidesn=Slides.Item(i);   
     Slidesn.Shapes.Range.Item(4).TextFrame.TextRange.Text=['Nr.','  ',STAND_TITLE{i}];
     Slidesn.Shapes.Range.Item(5).Table.Cell(1,2).Shape.TextFrame.TextRange.Text=PROJECT;   
     Slidesn.Shapes.Range.Item(5).Table.Cell(3,2).Shape.TextFrame.TextRange.Text=[num2str(floor(SOLL_WERT(i))),'N'];   
     Slidesn.Shapes.Range.Item(5).Table.Cell(4,2).Shape.TextFrame.TextRange.Text=[num2str(floor(IST_WERT(i))),'N'];  
     if IST_WERT(i)<SOLL_WERT(i)
      Slidesn.Shapes.Range.Item(5).Table.Cell(4,2).Shape.TextFrame.TextRange.Font.color.rgb=255;
     end
     
     sfilename=[filename,'MP' num2str(i) '.jpg'];
     Slidesn.Shapes.AddPicture(sfilename,1,1,357.5,275.5,396.1943,215.800);
     waitbar(i/length(MP));
 end
Presentation.Save
%%%%%%%%%%%%输出报告生成信息到公共空间%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='Audi保险杠拉力试验';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
close(t2);
msgbox('报告生成完毕');


% --------------------------------------------------------------------
function Menu1_Callback(hObject, eventdata, handles)
run Auto2_3_Config


% --------------------------------------------------------------------
function Menu2_Callback(hObject, eventdata, handles)
[filename,pathname,fileindex]=uigetfile('*.ppt;*.pptx','选择ppt');
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
    return
end
 if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
 end
filespec_user=fullfile(pathname,filename);
try

     Powerpoint = actxGetRunningServer('Powerpoint.Application');
 catch
     Powerpoint = actxserver('Powerpoint.Application'); 
 end
 %Powerpoint.Visible = 0;    


 if exist(filespec_user,'file')
     Presentation = Powerpoint.Presentation.Open(filespec_user);

 else
     Presentation = Powerpoint.Presentation.Add;
         Presentation.SaveAs(filespec_user);
 end
 

  Fileaddress=fullfile(pathname,'result');
   Slides = Powerpoint.ActivePresentation.Slides; 
  Slides_number=Slides.count;
  t1=waitbar(0,'正在生成图片');
  for i=1:Slides_number
 Slides1=Slides.Item(i);
  Slides1.Export([Fileaddress,'\',num2str(i),'.bmp'],'bmp');
   pic = imread([Fileaddress,'\',num2str(i),'.bmp']);
   pic_1 = imcrop(pic,[5.25,100,1010.5,612]);
   imwrite(pic_1,[Fileaddress,'\',num2str(i),'.bmp']);
   waitbar(i/(Slides_number));
  end
close(t1);
 try
       system('taskkill/IM POWERPNT.exe');
 end
       t2=waitbar(0,'正在生成Word报告');
 filespec_user=[Fileaddress,'\report.doc'];
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


InlineShapes=Document.InlineShapes;
He=180*1.26/8*9;
Wi=240*1.9/16.09*15;
for i=1:Slides_number
    sfilename1=[Fileaddress,'\',num2str(i),'.bmp'];
handle=Selection.InlineShapes.AddPicture(sfilename1);
InlineShapes.Item(i).Height=He;
InlineShapes.Item(i).Width=Wi;
delete(sfilename1); 
Selection.Start = Selection.end;
Selection.TypeParagraph;
waitbar(i/Slides_number);
end
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
winopen([Fileaddress,'\report.doc']);
close(t2);
