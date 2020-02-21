function varargout = Auto3_2_2(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto3_2_2_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto3_2_2_OutputFcn, ...
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


% --- Executes just before Auto3_2_2 is made visible.
function Auto3_2_2_OpeningFcn(hObject, eventdata, handles, varargin)
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

% UIWAIT makes Auto3_2_2 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto3_2_2_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on selection change in listbox2.
function listbox2_Callback(hObject, eventdata, handles)
cla(handles.axes1);
DATA_SCOPE=getappdata(0,'Auto3_2_2_DATA_SCOPE');
MP=getappdata(0,'Auto3_2_2_MP');
outmaxvalue=getappdata(0,'Auto3_2_2_OUT');
maxWegvalue=getappdata(0,'Auto3_2_2_maxWegvalue');
figuretitle=getappdata(0,'Auto3_2_2_figuretitle');
ZIHAO_TU_YULAN=10;
TITLEFONTSIZE=13;
Teilnr=get(handles.edit7, 'string');
CHOOSE=get(handles.listbox2,'Value');                %listbox的值
i=CHOOSE;
plot(handles.axes1,MP{3*i-2}(:,2),MP{3*i-2}(:,1),'linewidth',2);
hold on
plot(handles.axes1,MP{3*i-1}(:,2),MP{3*i-1}(:,1),'Color','r','linewidth',2);
plot(handles.axes1,MP{3*i}(:,2),MP{3*i}(:,1),'Color','g','linewidth',2);

xlabel(handles.axes1,'Weg/位移[mm]','FontSize',ZIHAO_TU_YULAN)
ylabel(handles.axes1,'Kraft/力[N]','FontSize',ZIHAO_TU_YULAN)
title(handles.axes1,figuretitle{i},'FontSize',TITLEFONTSIZE)
axis(handles.axes1,[0 max([maxWegvalue(i*3-2),maxWegvalue(i*3-1),maxWegvalue(i*3)])*1.05 0 1.1*max([outmaxvalue(i*3-2),outmaxvalue(i*3-1),outmaxvalue(i*3)])]);
legend(handles.axes1,'1#','2#','3#');
set(handles.edit1,'String',num2str(outmaxvalue(i*3-2),'%.1f'));
set(handles.edit2,'String',num2str(outmaxvalue(i*3-1),'%.1f'));
set(handles.edit3,'String',num2str(outmaxvalue(i*3),'%.1f'));
  
  
function listbox2_CreateFcn(hObject, eventdata, handles)

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



function edit3_Callback(hObject, eventdata, handles)

function edit3_CreateFcn(hObject, eventdata, handles)

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


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
if isempty(get(handles.edit7, 'string'))
        msgbox('请输入零件号');
      return;
end
DATA_TYPE_KRAFT=get(handles.popupmenu3,'value');      %读取数据第几列为力
DATA_TYPE_WEG=get(handles.popupmenu4,'value');          %读取数据第几列为位移
DATA_SCOPE=get(handles.popupmenu5,'value');     %读取数据范围
DATA_SCOPE_STRING_list=get(handles.popupmenu5,'string'); 
DATA_SCOPE_STRING=DATA_SCOPE_STRING_list{DATA_SCOPE};
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据','MultiSelect','on');
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
    return;
else
     Teilnr=get(handles.edit7, 'string');
    if DATA_SCOPE==4
        [a b]=size(filename)
        if b~=36
            msgbox('请一次性导入该零件号的所有36组数据');
            return;
        else
            t1=waitbar(0,'正在读入数据');
            for i=1:length(filename)
                Filename{i}=strcat(pathname,filename{i});
                [Type Sheet Format]=xlsfinfo(Filename{i}) ;
                sheet{i}=Sheet;
                MP_Mittle{i}=xlsread(Filename{i},char(sheet{1,i}(1,4)));
                MP{i}(:,2)=MP_Mittle{i}(:,DATA_TYPE_WEG);
                MP{i}(:,1)=MP_Mittle{i}(:,DATA_TYPE_KRAFT);
                waitbar(i/length(filename));
                try
                    system('taskkill/IM excel.exe');
                end
            end
            listtitle={'端头1--40℃','端头1-80℃','端头1-RT','端头2--40℃','端头2-80℃','端头2-RT','端头3--40℃','端头3-80℃','端头3-RT','端头4--40℃','端头4-80℃','端头4-RT'};
        end
    else
        [a b]=size(filename)
        if b~=12
            msgbox('请一次性导入该零件号的所有12组数据');
            return;
        else
            t1=waitbar(0,'正在读入数据');
            for i=1:length(filename)
                Filename{i}=strcat(pathname,filename{i});
                [Type Sheet Format]=xlsfinfo(Filename{i}) ;
                sheet{i}=Sheet;
                MP_Mittle{i}=xlsread(Filename{i},char(sheet{1,i}(1,4)));
                MP{i}(:,2)=MP_Mittle{i}(:,DATA_TYPE_WEG);
                MP{i}(:,1)=MP_Mittle{i}(:,DATA_TYPE_KRAFT);
                waitbar(i/length(filename));
                try
                    system('taskkill/IM excel.exe');
                end
            end
            listtitle={['端头1-',DATA_SCOPE_STRING],['端头2-',DATA_SCOPE_STRING],['端头3-',DATA_SCOPE_STRING],['端头4-',DATA_SCOPE_STRING]};
        end
    end
end
set(handles.listbox2,'String',listtitle);
close(t1);
if DATA_SCOPE==4
    for i=1:36
        outmaxvalue(i,1)=max(MP{i}(:,1));
        maxWegvalue(i,1)=max(MP{i}(:,2));
    end
    figuretitle={[Teilnr,' bei -40°C Endstueck-1'],[Teilnr,' bei 80°C Endstueck-1'],[Teilnr,' bei RT Endstueck-1'],...
        [Teilnr,' bei -40°C Endstueck-2'],[Teilnr,' bei 80°C Endstueck-2'],[Teilnr,' bei RT Endstueck-2'],...
        [Teilnr,' bei -40°C Endstueck-3'],[Teilnr,' bei 80°C Endstueck-3'],[Teilnr,' bei RT Endstueck-3'],...
        [Teilnr,' bei -40°C Endstueck-4'],[Teilnr,' bei 80°C Endstueck-4'],[Teilnr,' bei RT Endstueck-4']};
else
       for i=1:12
        outmaxvalue(i,1)=max(MP{i}(:,1));
        maxWegvalue(i,1)=max(MP{i}(:,2));
    end
    figuretitle={[Teilnr,' bei ', DATA_SCOPE_STRING ' Endstueck-1'],[Teilnr,' bei ', DATA_SCOPE_STRING ' Endstueck-2'],...
        [Teilnr,' bei ', DATA_SCOPE_STRING ' Endstueck-3'],[Teilnr,' bei ', DATA_SCOPE_STRING ' Endstueck-4']};
end
setappdata(0,'Auto3_2_2_pathname',pathname);
setappdata(0,'Auto3_2_2_filename',filename);
setappdata(0,'Auto3_2_2_MP',MP);
setappdata(0,'Auto3_2_2_OUT',outmaxvalue);
setappdata(0,'Auto3_2_2_maxWegvalue',maxWegvalue);
setappdata(0,'Auto3_2_2_figuretitle',figuretitle);
setappdata(0,'Auto3_2_2_DATA_SCOPE',DATA_SCOPE);
setappdata(0,'Auto3_2_2_DATA_SCOPE_STRING',DATA_SCOPE_STRING);
set(handles.pushbutton2,'Enable','on');


% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
pathname=getappdata(0,'Auto3_2_2_pathname');
filename=getappdata(0,'Auto3_2_2_filename');
Fileadress=strcat(pathname,'result\');
   if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
   end
MP=getappdata(0,'Auto3_2_2_MP');
outmaxvalue=getappdata(0,'Auto3_2_2_OUT');
maxWegvalue=getappdata(0,'Auto3_2_2_maxWegvalue');
figuretitle=getappdata(0,'Auto3_2_2_figuretitle');
DATA_SCOPE=getappdata(0,'Auto3_2_2_DATA_SCOPE');
DATA_SCOPE_STRING=getappdata(0,'Auto3_2_2_DATA_SCOPE_STRING');
ZIHAO_TU_YULAN=20;
TITLEFONTSIZE=30;
Teilnr=get(handles.edit7, 'string');
t1=waitbar(0,'正在生成图片') ;  
if DATA_SCOPE==4
    Figurenum=12;
else
    Figurenum=4;
end
 for i=1: Figurenum
    h(i)=figure;
    set(h(i),'position',[100,100,1300,800]); 
    set(h(i),'visible','off');    
    plot(MP{3*i-2}(:,2),MP{3*i-2}(:,1),'linewidth',2);
    hold on
    plot(MP{3*i-1}(:,2),MP{3*i-1}(:,1),'Color','r','linewidth',2);
    plot(MP{3*i}(:,2),MP{3*i}(:,1),'Color','g','linewidth',2);
    grid on;
      set(gca,'FontSize',ZIHAO_TU_YULAN)
    xlabel('Weg/位移[mm]','FontSize',ZIHAO_TU_YULAN)
    ylabel('Kraft/力[N]','FontSize',ZIHAO_TU_YULAN)
    title(figuretitle{i},'FontSize',TITLEFONTSIZE)
    axis([0 max([maxWegvalue(i*3-2),maxWegvalue(i*3-1),maxWegvalue(i*3)])*1.05 0 1.1*max([outmaxvalue(i*3-2),outmaxvalue(i*3-1),outmaxvalue(i*3)])]);
    legend('1#','2#','3#');
     sfilename1=[Fileadress,num2str(i),'-',figuretitle{i},'.jpg'];
    saveas(h(i),sfilename1);
    close(h(i));
    waitbar(i/Figurenum);
end
close(t1);
t2=waitbar(0,'正在生成Word报告') ;
biaotihao=10;
He=180*1.0771653543307086614173228346457;
Wi=240*1.9;
filespec_user=[Fileadress,[get(handles.edit7,'String'),'-report.doc']];
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
waitbar(0.1);
headline='III.1 TIB Bowdenzug Abzugsversuch 内开拉丝端头拉脱力试验';
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=biaotihao; % 字体大小
Content.Font.NameAscii='Arial';
Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落         

  
lc=28.381133333333333333333333333333; %厘米换算
column_width = [3*lc,3.25*lc,3.25*lc];
for i=1:Figurenum
Teiladdress{i}=[Fileadress,num2str(i),'-',figuretitle{i},'.jpg'];
end

if DATA_SCOPE==4    
    for z=1:4
        paptitle={'Endstueck-1/端头1','Endstueck-2/端头2','Endstueck-3/端头3','Endstueck-4/端头4'}    ;
        headline=paptitle{z};
        Selection.Text=headline;
        Selection.Font.NameAscii='Arial';
        Selection.Font.Size=biaotihao; % 字体大小
        Selection.Start=Selection.end;
        Selection.Start = Content.end;
        Selection.TypeParagraph;% 插入一个新的空段落        
        Tab = Document.Tables.Add(Selection.Range, 10, 3);
        DTI = Document.Tables.Item(z); % 表格句柄
        DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
        DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
        % 设置行高，列宽        
        for i = 1:3
            DTI.Columns.Item(i).Width = column_width(i);
        end
        for i=1: 10
            for j=1:3
                DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
                DTI.Cell(i,j).Range.Font.NameAscii='Arial';
                DTI.Cell(i,j).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
            end
        end
        DTI.Cell(1,1).Merge(DTI.Cell(1,2));
        DTI.Cell(2,1).Merge(DTI.Cell(4,1));
        DTI.Cell(5,1).Merge(DTI.Cell(7,1));
        DTI.Cell(8,1).Merge(DTI.Cell(10,1));
        DTI.Cell(1,1).Range.Text = 'Prufling试验件';
        DTI.Cell(1,2).Range.Text = '拉脱力Abzugskraft(N)';
        DTI.Cell(2,1).Range.Text = 'Bei -40℃';
        DTI.Cell(5,1).Range.Text = 'Bei 80℃';
        DTI.Cell(8,1).Range.Text = 'Bei RT';
        DTI.Cell(2,2).Range.Text = '1#';
        DTI.Cell(3,2).Range.Text = '2#';
        DTI.Cell(4,2).Range.Text = '3#';
        DTI.Cell(5,2).Range.Text = '1#';
        DTI.Cell(6,2).Range.Text = '2#';
        DTI.Cell(7,2).Range.Text = '3#';
        DTI.Cell(8,2).Range.Text = '1#';
        DTI.Cell(9,2).Range.Text = '2#';
        DTI.Cell(10,2).Range.Text = '3#';
        for i=2:10
            DTI.Cell(i,3).Range.Text =num2str(outmaxvalue(i-1+(z-1)*9),'%.1f');
        end        
        Selection.Start = Content.end;
        Selection.TypeParagraph;
        InlineShapes=Document.InlineShapes;        
        for i=1:3
            handle=Selection.InlineShapes.AddPicture(Teiladdress{1,i+(z-1)*3});
            InlineShapes.Item(i+(z-1)*3).Height=He;
            InlineShapes.Item(i+(z-1)*3).Width=Wi;
            Selection.Start = Selection.end;
            Selection.TypeParagraph;% 插入一个新的空段落
            Selection.Start = Selection.end;
            Selection.TypeParagraph;% 插入一个新的空段落
        end
        waitbar(0.1+0.2*z);
    end
else
     for z=1:4
        paptitle={'Endstueck-1/端头1','Endstueck-2/端头2','Endstueck-3/端头3','Endstueck-4/端头4'}    ;
        headline=paptitle{z};
        Selection.Text=headline;
        Selection.Font.NameAscii='Arial';
        Selection.Font.Size=biaotihao; % 字体大小
        Selection.Start=Selection.end;
        Selection.Start = Content.end;
        Selection.TypeParagraph;% 插入一个新的空段落        
        Tab = Document.Tables.Add(Selection.Range, 4, 3);
        DTI = Document.Tables.Item(z); % 表格句柄
        DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
        DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
        % 设置行高，列宽        
        for i = 1:3
            DTI.Columns.Item(i).Width = column_width(i);
        end
        for i=1: 4
            for j=1:3
                DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
                DTI.Cell(i,j).Range.Font.NameAscii='Arial';
                DTI.Cell(i,j).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
            end
        end
        DTI.Cell(1,1).Merge(DTI.Cell(1,2));
        DTI.Cell(2,1).Merge(DTI.Cell(4,1));   
        DTI.Cell(1,1).Range.Text = 'Prufling试验件';
        DTI.Cell(1,2).Range.Text = '拉脱力Abzugskraft(N)';
        DTI.Cell(2,1).Range.Text = ['Bei ',DATA_SCOPE_STRING];
        DTI.Cell(2,2).Range.Text = '1#';
        DTI.Cell(3,2).Range.Text = '2#';
        DTI.Cell(4,2).Range.Text = '3#';     
        for i=2:4
            DTI.Cell(i,3).Range.Text =num2str(outmaxvalue(i-1+(z-1)*3),'%.1f');
        end        
        Selection.Start = Content.end;
        Selection.TypeParagraph;
        InlineShapes=Document.InlineShapes;
        
        handle=Selection.InlineShapes.AddPicture(Teiladdress{1,z});
        InlineShapes.Item(z).Height=He;
        InlineShapes.Item(z).Width=Wi;
        Selection.Start = Selection.end;
        Selection.TypeParagraph;% 插入一个新的空段落
        Selection.Start = Selection.end;
        Selection.TypeParagraph;% 插入一个新的空段落        
        waitbar(0.1+0.2*z);
    end
end

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
for i=1:Figurenum
    delete(Teiladdress{1,i});
end
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='内开拉丝端头拉脱力';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
waitbar(1);
close(t2);
winopen(filespec_user);


function edit7_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function edit7_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu3.
function popupmenu3_Callback(hObject, eventdata, handles)

% --- Executes during object creation, after setting all properties.
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


% --- Executes on selection change in popupmenu5.
function popupmenu5_Callback(hObject, eventdata, handles)

function popupmenu5_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
