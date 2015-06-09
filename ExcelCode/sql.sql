

create function [dbo].[getspecialnotes](@sitnum nvarchar(20))
returns nvarchar(max)
AS
begin

declare @notes nvarchar(max)

select @notes=notes from
(
select SitNum,Notes from dbo.RevData 
union select SitNum,Bunch from dbo.Bunchment where  SitNum in(select SitNum from dbo.TempStudentInYear )
union select Sitnum,convert_resson from dbo.converts 
union select SitNum,'إيقاف قيد' from dbo.KaidData  where IsWork=0  ) a
where sitnum=@sitnum
return @notes
end

GO


ALTER PROCEDURE [dbo].[GetAllStdDegByClass]
	-- Add the parameters for the stored procedure here
	(@ClassId as int) 
	
AS
BEGIN
declare @YearId int set @YearId=(select YearId from dbo.StudyYear where IsCurrent='True' )
select dbo.GetNum(dbo.StudentsData.SitNum,@YearId) as Num,dbo.StudentsData.SitNum,StdName,Des,
dbo.getirregular(dbo.StudentsData.SitNum) as IsIrregular,SecrtNum,dbo.TempStudentsDegrees.SubjId,SubjName,(select top 1 BrClassId from dbo.BrClassSubjects where SubjId=dbo.TempStudentsDegrees.SubjId) as SubjYearId, dbo.GetSubjClass(dbo.TempStudentsDegrees.SubjId)SubjYName,dbo.GetStdSubjBr(dbo.TempStudentsDegrees.SubjId,0,dbo.StudentsData.SitNum,@YearId) As Remaning,
dbo.GetZero(dbo.TempStudentsDegrees.OralDeg)OralDeg,dbo.GetZero(dbo.TempStudentsDegrees.WriringDeg)WriringDeg,dbo.GetZero(SumSubjDeg) as Total, dbo.GetZero(dbo.GetDeg(SumSubjDeg,HelpDegOnSubj))as LastTotal,  HelpDegOnSubj,dbo.ChangeGradeInk(dbo.GetGradeName(SumSubjDeg,dbo.TempStudentsDegrees.SubjId))as Grade, (select case dbo.ChangeGradeInk(dbo.GetGradeName(SumSubjDeg,dbo.TempStudentsDegrees.SubjId)) when 'غ' then 'غ' else dbo.ChangeGradeInk(GradeName) end) as LastGrade, (select dbo.GetZero(dbo.GetDeg(StdTotalDeg,HelpDegOnTotalDeg)) as StdTotalDeg from dbo.TempStudentInYear where YearId=@YearId and SitNum=dbo.TempStudentsDegrees.SitNum)as TotalDeg, (select dbo.GetZero(StdTotalDeg) from dbo.TempStudentInYear where YearId=@YearId and SitNum=dbo.TempStudentsDegrees.SitNum)as TotalBefore, (select HelpDegOnTotalDeg from dbo.TempStudentInYear where YearId=@YearId and SitNum=dbo.TempStudentsDegrees.SitNum)as HelpDegOnTotalDeg,IsFromLastYear,subjectState, (select StdState from dbo.TempStudentInYear where YearId=@YearId and SitNum=dbo.TempStudentsDegrees.SitNum)as StdState, (select IsFinal from dbo.TempStudentInYear where YearId=@YearId and SitNum=dbo.TempStudentsDegrees.SitNum)as IsFinal, (select dbo.GetGrade(StdGradeId) from dbo.TempStudentInYear where YearId=@YearId and SitNum=dbo.TempStudentsDegrees.SitNum)as StdGrade,  
dbo.GetGrade(dbo.GetTotalGradeBefore(dbo.TempStudentInYear.StdTotalDeg,dbo.TempStudentInYear.BrClassId)) as TotalGradeBefore,
(select Term from dbo.Subjects where SubjId=dbo.TempStudentsDegrees.SubjId) as SubjTerm,Subjects.OralDeg Oral,MaxStayDes.Id MaxStayId,
(select MaxTotal from dbo.MaxTotal where BrId=dbo.TempStudentInYear.BrClassId)/2 as HalfMaxTotal,
dbo.getspecialnotes(dbo.StudentsData.SitNum) as specialnotes
from dbo.TempStudentsDegrees join dbo.StudentsData on dbo.TempStudentsDegrees.SitNum = dbo.StudentsData.SitNum  join dbo.Subjects on dbo.TempStudentsDegrees.SubjId=dbo.Subjects.SubjId join dbo.Grades on dbo.TempStudentsDegrees.GradeId=dbo.Grades.GradeId  join dbo.TempStudentInYear on dbo.TempStudentsDegrees.SitNum=dbo.TempStudentInYear.SitNum and dbo.TempStudentsDegrees.YearId=dbo.TempStudentInYear.YearId  join MaxStayDes on dbo.TempStudentInYear.Position_Id=MaxStayDes.Id  join SecretNum on dbo.TempStudentsDegrees.SitNum=dbo.SecretNum.SitNum and dbo.TempStudentsDegrees.YearId=SecretNum.YearId where dbo.TempStudentsDegrees.YearId=@YearId and dbo.TempStudentsDegrees.SitNum in(select SitNum from dbo.TempStudentInYear where BrClassId=@ClassId and YearId=@YearId)

/*إذا أزلت السطر التالي ستظهر نتيجة مواد التخلف*/
/*and dbo.TempStudentsDegrees.SubjId in(select SubjId from dbo.BrClassSubjects where BrClassId=@ClassId)*/

/*الكود التالي لجلب بيانات مواد التخلف لفرقة بعينها*/
/*declare @YearId int  set @YearId=(select YearId from dbo.StudyYear where IsCurrent='True' )
select dbo.GetNum(dbo.TempStudentsDegrees.SitNum,@YearId) as Num, dbo.TempStudentsDegrees.SitNum,dbo.TempStudentsDegrees.SubjId,SubjName, dbo.GetSubjClass(dbo.TempStudentsDegrees.SubjId)SubjYear,dbo.GetZero(dbo.TempStudentsDegrees.OralDeg)OralDeg,dbo.GetZero(dbo.TempStudentsDegrees.WriringDeg)WriringDeg,dbo.GetZero(SumSubjDeg) as Total,HelpDegOnSubj,dbo.GetZero(dbo.GetDeg(SumSubjDeg,HelpDegOnSubj)) as LastTotal, dbo.ChangeGradeInk(dbo.GetGradeName(SumSubjDeg, dbo.TempStudentsDegrees.SubjId))as Grade,(select case dbo.ChangeGradeInk(dbo.GetGradeName(SumSubjDeg,dbo.TempStudentsDegrees.SubjId)) when 'غ' then 'غ' else dbo.ChangeGradeInk(GradeName) end) as LastGrade,subjectState
 from dbo.TempStudentsDegrees join dbo.TempStudentsData on dbo.TempStudentsDegrees.SitNum = dbo.TempStudentsData.SitNum
join dbo.Subjects on dbo.TempStudentsDegrees.SubjId=dbo.Subjects.SubjId join dbo.Grades on dbo.TempStudentsDegrees.GradeId=dbo.Grades.GradeId
join dbo.TempStudentInYear on dbo.TempStudentsDegrees.SitNum=dbo.TempStudentInYear.SitNum and dbo.TempStudentsDegrees.YearId=dbo.TempStudentInYear.YearId
where dbo.TempStudentsDegrees.YearId=@YearId and dbo.TempStudentsDegrees.SitNum in(select SitNum from dbo.TempStudentInYear where BrClassId=@ClassId and YearId=@YearId)
and dbo.TempStudentsDegrees.SubjId not in(select SubjId from dbo.BrClassSubjects where BrClassId=@ClassId)*/

/*
Id's For Classes
1= أولى أصول
2= ثانية أصول
3= ثالثة أصول قسم التفسير
4= رابعة أصول قسم التفسير
5= ثالثة أصول قسم الحديث
6= رابعة أصول قسم الحديث
7= ثالثة أصول قسم العقيدة
8= رابعة أصول قسم العقيدة
9= أولى شريعة
10= ثانية شريعة
11= ثالثة شريعة
12= رابعة شريعة
13= رابعة أصول قسم التفسير - قديم
14= رابعة أصول قسم الحديث - قديم
15= رابعة أصول قسم العقيدة - قديم
16= رابعة شريعة - قديم
*/
END

GO




