%%

clear all;
close all;

%% user options
employeeName="all"; % cognome persona, o "ricerca"/"research", o "operatori", o "tutti"/"all" 
masterExcelShifts="P:\Turni Macchina\Turni_Gennaio-Giugno2026_ver1.xlsx";
omFileName=ExtractFileName(masterExcelShifts);
% tMin=datetime(now,'ConvertFrom','datenum'); % tMin=datetime('18/10/2025 13:05:00','InputFormat','dd/MM/uuuu HH:mm:ss');
tMax=datetime('26/06/2026 23:59:59','InputFormat','dd/MM/uuuu HH:mm:ss');

%% parsing original excel
masterShifts=ParseMasterFile(masterExcelShifts);

%% crunch comments
% fprintf("checking that FM is never shift leader...\n")
% find(contains(lower(masterShifts.(8)),"recupero"))
if (isstring(employeeName))
    if (strcmpi(employeeName,"all") | strcmpi(employeeName,"tutti") )
        employeeNames=GetUniqueNames(masterShifts);
    elseif (contains(lower(employeeName),"ricerca") | contains(lower(employeeName),"research"))
        employeeNames=["Butella","Donetti","Felcini","Mereghetti","Pullia","Ricerca","Savazzi","Nodari"];
        employeeNickNames=["GB","MD","EF","AM","MGP",missing(),"SS","MN"];
    elseif (contains(lower(employeeName),"operatori") | contains(lower(employeeName),"operators"))
        employeeNames=["Abbiati","Basso","Beretta","Bozza","Cappello","Chiesa","Liceti","Malinverni","Manzini","Scotti","Spairani"];
    else
        employeeNames=[employeeName];
    end
else
    employeeNames=employeeName;
end

%% process
% - time window
if (~exist("tMin","var")), tMin=missing(); end
if (ismissing(tMin)), tMin=min(masterShifts.Turni); end
if (~exist("tMax","var")), tMax=missing(); end
if (ismissing(tMax)), tMax=max(masterShifts.Turni); end
crunchShifts=masterShifts(isbetween(masterShifts.Turni,tMin,tMax),:);
% - main loop
rEmployeeShifts=NaN();
for ii=1:length(employeeNames)
    % - build table of employee
    employeeShifts=BuildEmployeeTable(crunchShifts,employeeNames(ii));
    % - write csv file
    if (contains(lower(employeeName),"ricerca"))
        NickName=employeeNickNames(ii);
        if (~ismissing(NickName)), employeeShifts.subjects(:)=strcat(NickName,": ",employeeShifts.subjects(:));  end
        if (ii==1)
            rEmployeeShifts=employeeShifts;
        else
            rEmployeeShifts=[rEmployeeShifts;employeeShifts];
        end
    else
        oFileName=sprintf("%s_%s.csv",employeeNames(ii),omFileName);
        writeGoogleCalendarCSV(employeeShifts,oFileName);
    end
end
if (contains(lower(employeeName),"ricerca"))
    oFileName=sprintf("%s_%s.csv",employeeName,omFileName);
    writeGoogleCalendarCSV(rEmployeeShifts,oFileName);
end

%%

%% functions
function masterShifts=ParseMasterFile(masterExcelShifts)
    shiftNames=["morning shift","afternoon shift","night shift"];
    %
    fprintf("parsing master file %s ...\n",masterExcelShifts);
    masterShifts=readtable(masterExcelShifts);
    fprintf("...done: acquired %d lines and %d columns;\n",size(masterShifts,1),size(masterShifts,2));
    %
    fprintf("checking that FM is never shift leader...\n")
    for iCol=2:2:6
        indices=find(strcmpi(masterShifts.(iCol),"FM"));
        if (~isempty(indices))
            fprintf("...found it in %s %d times!\n",shiftNames(iCol/2),length(indices));
            [masterShifts.(iCol)(indices),masterShifts.(iCol+1)(indices)]=deal(masterShifts.(iCol+1)(indices),masterShifts.(iCol)(indices));
        end
    end
    fprintf("...done;\n");
    %
    fprintf("checking cancelled shifts...\n");
    Excel = actxserver('Excel.Application');
    Excel.Workbooks.Open(masterExcelShifts);
    for iRow=1:size(masterShifts,1)
        for iCol=1:6
            Range= Excel.Range(sprintf("%s%d",char(double('A')+iCol),iRow+1));
            if (Range.Font.Strikethrough)
                if (strlength(string(masterShifts.(iCol+1)(iRow)))>0)
                    fprintf("...cacelled %s shift for %s on %s!\n",shiftNames(ceil(iCol/2)),string(masterShifts.(iCol+1)(iRow)),string(masterShifts.(1)(iRow)));
                end
                masterShifts.(iCol+1)(iRow)=cellstr("");
            end
        end
    end
    Quit(Excel);
    delete(Excel);
    fprintf("...done;\n");
end

function oNames=CapitalizeNames(iNames)
    oNames=lower(iNames);
    iFMs=strcmp(oNames,"fm");
    oNames(iFMs)="FM";
    oNames(strlength(oNames)==0)="Nessuno";
    oNames(~iFMs)=compose("%s%s",upper(extractBetween(oNames(~iFMs),1,1)),extractBetween(oNames(~iFMs),2,strlength(oNames(~iFMs))));
end

function employeeShifts=BuildEmployeeTable(masterShifts,employeeName)
    % preamble
    shiftHours=["06:00","14:00","14:00","22:00","22:00","06:00"];
    shiftDays=zeros(1,6); shiftDays(5:6)=1;
    shiftRoles=["shift leader","addetto sicurezza"];
    shiftTag=["turno mattino","turno pomeriggio","turno notte"];
    % do the job
    fprintf("looking for shifts of %s...\n",employeeName);
    employeeShifts=table();
    for iCol=2:7
        iTurni=contains(masterShifts.(iCol),employeeName,"IgnoreCase",true);
        nTurni=sum(iTurni);
        if (nTurni>0)
            % prepare info to store
            currDates=masterShifts.(1);
            currLen=size(employeeShifts,1);
            otherShifters=masterShifts.(iCol+(-1)^mod(iCol,2)); otherShifters=CapitalizeNames(string(otherShifters(iTurni)));
            % - subject
            employeeShifts.subjects(currLen+1:currLen+nTurni)=shiftTag(floor(iCol/2));
            % - start dates and times:
            employeeShifts.startDates(currLen+1:currLen+nTurni)=currDates(iTurni);
            employeeShifts.startTimes(currLen+1:currLen+nTurni)=shiftHours(2*floor(iCol/2)-1);
            % - end dates and times:
            if (2*floor(iCol/2)==length(shiftHours))
                % shift ends the following day
                currEndDates=datenum(currDates(iTurni));
                for ii=1:length(currEndDates)
                    currEndDates(ii)=addtodate(currEndDates(ii),shiftDays(iCol-1),"day");
                end
                employeeShifts.endDates(currLen+1:currLen+nTurni)=datetime(currEndDates,"ConvertFrom","datenum");
            else
                employeeShifts.endDates(currLen+1:currLen+nTurni)=currDates(iTurni);
            end
            employeeShifts.endTimes(currLen+1:currLen+nTurni)=shiftHours(2*floor(iCol/2));
            % - descriptions
%             employeeShifts.descriptions(currLen+1:currLen+nTurni)=compose("Tu sei %s, con %s come %s",...
%                     shiftRoles(mod(iCol,2)+1),otherShifters,shiftRoles(mod(iCol-1,2)+1));
            employeeShifts.descriptions(currLen+1:currLen+nTurni)=compose("Sei in turno con %s",otherShifters);
            % - take into account notes in same column
            additionalNotes=string(masterShifts.(iCol));
            additionalNotes=additionalNotes(iTurni);
            iNotes=find(~strcmpi(additionalNotes,employeeName));
            employeeShifts.descriptions(currLen+iNotes)=additionalNotes(iNotes);
            % - take into account notes in columns >7
            additionalComments=string(masterShifts.(floor(iCol/2)+7));
            additionalComments=additionalComments(iTurni);
            iComments=find(strlength(additionalComments)>0);
            employeeShifts.descriptions(currLen+iComments)=compose("%s;\n NOTA: %s",employeeShifts.descriptions(currLen+iComments),additionalComments(iComments));
            % - take into accounts possible time indications
            shiftCol=string(masterShifts.(iCol)); shiftCol=shiftCol(iTurni);
            if (nTurni==1)
                iShifts=~isempty(regexp(shiftCol,"9\s*-\s*17")) | ~isempty(regexp(additionalComments,"9\s*-\s*17"));
                if (iShifts), employeeShifts.startTimes(currLen+iShifts)="09:00"; employeeShifts.endTimes(currLen+iShifts)="17:00"; clear iShifts; end
                iShifts=~isempty(regexp(shiftCol,"8\s*-\s*16")) | ~isempty(regexp(additionalComments,"8\s*-\s*16"));
                if (iShifts), employeeShifts.startTimes(currLen+iShifts)="08:00"; employeeShifts.endTimes(currLen+iShifts)="16:00"; clear iShifts; end
            else
                iShifts=find(~cellfun(@isempty,regexp(shiftCol,"9\s*-\s*17")) | ~cellfun(@isempty,regexp(additionalComments,"9\s*-\s*17")));
                employeeShifts.startTimes(currLen+iShifts)="09:00"; employeeShifts.endTimes(currLen+iShifts)="17:00"; clear iShifts;
                iShifts=find(~cellfun(@isempty,regexp(shiftCol,"8\s*-\s*16")) | ~cellfun(@isempty,regexp(additionalComments,"8\s*-\s*16")));
                employeeShifts.startTimes(currLen+iShifts)="08:00"; employeeShifts.endTimes(currLen+iShifts)="16:00"; clear iShifts;
            end
        end
    end
    employeeShifts.descriptions(contains(employeeShifts.descriptions,"nessuno","IgnoreCase",true))="Sei in turno da solo";
    % employeeShifts.descriptions(contains(employeeShifts.descriptions,"ricerca","IgnoreCase",true))="Turno ricerca";
    % - sorting
    [~,IDs]=sort(employeeShifts.startDates);
    employeeShifts=employeeShifts(IDs,:);
    %
    fprintf("...found %d shifts!\n",size(employeeShifts,1));
end

function writeGoogleCalendarCSV(employeeShifts,oFileName)
    % - formatting
    employeeShifts.startDates.Format='yyyy-MM-dd';
    employeeShifts.endDates.Format='yyyy-MM-dd';
    headers=["Start Date","Start Time","End Date","End Time","Subject","Description"];
    % - do the actual job
    fprintf("preparing %s file for import into google calendar...\n",oFileName);
    nDataRows=size(employeeShifts,1);
    employTable=strings(nDataRows+1,length(headers));
    employTable(1,:)=headers;
    for ii=1:length(headers)
        switch headers(ii)
            case "Start Date"
                employTable(2:end,ii)=string(employeeShifts.startDates);
            case "Start Time"
                employTable(2:end,ii)=employeeShifts.startTimes;
            case "End Date"
                employTable(2:end,ii)=string(employeeShifts.endDates);
            case "End Time"
                employTable(2:end,ii)=employeeShifts.endTimes;
            case "Subject"
                employTable(2:end,ii)=employeeShifts.subjects;
            case "Description"
                employTable(2:end,ii)=employeeShifts.descriptions;
            otherwise
                error("Unknown header: %s",headers(ii));
        end
    end
    writematrix(employTable,oFileName,"QuoteStrings",true);
    fprintf("...done;\n");
end

function uniqueNames=GetUniqueNames(masterShifts)
    uniqueNames=unique(lower(string(masterShifts{:,2:7})));
    % handling of exceptions
    uniqueNames=replace(uniqueNames,"/",",");       % split names separated by "/"
    uniqueNames=replace(uniqueNames,"checklist","");% remove "checklist"
    uniqueNames=replace(uniqueNames,"chiusura",""); % remove "chiusura"
    uniqueNames=replace(uniqueNames,"turno","");    % remove "turno"
    uniqueNames=replace(uniqueNames,"8-16","");     % remove "8-16"
    uniqueNames=replace(uniqueNames,"9-17","");     % remove "9-17"
    uniqueNames=replace(uniqueNames,"*","");        % remove "*"
    uniqueNames=regexprep(uniqueNames,'\s+',',');   % consecutive empty spaces to ","
    uniqueNames=regexprep(uniqueNames,':+','');     % consecutive ":" to ""
    uniqueNames=regexprep(uniqueNames,',+',',');    % consecutive "," to single ","
    % extract entries with commas:
    tmp=uniqueNames(contains(uniqueNames,","));
    uniqueNames(contains(uniqueNames,","))=[];
    for ii=1:length(tmp)
        uniqueNames=[uniqueNames;split(tmp(ii),",")];
    end
    % remove IDs of operators
    for ii=0:9
        uniqueNames(strcmpi(uniqueNames,sprintf("%d",ii)))=[];
    end
    % remove odd strings
    oddStrings=["MINI","FERMO","AT"];
    for ii=1:length(oddStrings)
        uniqueNames(strcmpi(uniqueNames,oddStrings(ii)))=[];
    end
    % final settings
    uniqueNames=unique(uniqueNames);
    uniqueNames=uniqueNames(strlength(uniqueNames)>0);
end

function oFileName=ExtractFileName(masterExcelShifts)
    [folder,baseFileNameNoExt,extension]=fileparts(masterExcelShifts);
    oFileName=regexprep(baseFileNameNoExt,'\s+','_'); % replace consecutive empty spaces to "_"
end
