using System.Collections.Generic;
using System.Web;
using System.Web.Mvc;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.Owin;
using MvcSpAppWeb.Models;
using MvcSpAppWeb.CodeHelper;
using MvcSpAppWeb.Config;

namespace MvcSpAppWeb.Controllers
{
    [Authorize]
    public class SpListInfoController : Controller
    {
        private ApplicationSignInManager _signInManager;
        private ApplicationUserManager _userManager;

        public ApplicationSignInManager SignInManager
        {
            get
            {
                return _signInManager ?? HttpContext.GetOwinContext().Get<ApplicationSignInManager>();
            }
            private set
            {
                _signInManager = value;
            }
        }

        public ApplicationUserManager UserManager
        {
            get
            {
                return _userManager ?? HttpContext.GetOwinContext().GetUserManager<ApplicationUserManager>();
            }
            private set
            {
                _userManager = value;
            }
        }

        // GET: SpListInfo
        public ActionResult Index()
        {
            List<SpListMetadataModel> listModel = SpHelper.getAllSpLists(UserManager.FindById(User.Identity.GetUserId()), UserManager, Configuration.hostwebUrl);
            return View(listModel);
        }

        // GET: SpListInfo/Details/5
        public ActionResult Details(string id)
        {
            List<Dictionary<string, string>> spListItemsMetadatas = SpHelper.getSpListById(UserManager.FindById(User.Identity.GetUserId()), UserManager, id, Configuration.hostwebUrl);
            List<string> itemsList = new List<string>();
            #region Put All Items in one Single List
            int itemFieldsCount = 0;
            for (int i = 0; i < spListItemsMetadatas.Count; i++)
            {
                Dictionary<string, string> currendico = spListItemsMetadatas[i];
                if (i == 0)
                {
                    itemFieldsCount = currendico.Count;

                    foreach (var item in currendico)
                    {
                        itemsList.Add(item.Key);
                    }

                }
                foreach (var item in currendico)
                {
                    itemsList.Add(item.Value);
                }

            }
            #endregion
            ViewBag.itemFieldsCount = itemFieldsCount;
            return View(itemsList);
        }

        // GET: SpListInfo/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: SpListInfo/Create
        [HttpPost]
        public ActionResult Create(FormCollection collection)
        {
            try
            {
                // TODO: Add insert logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: SpListInfo/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: SpListInfo/Edit/5
        [HttpPost]
        public ActionResult Edit(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add update logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }

        // GET: SpListInfo/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: SpListInfo/Delete/5
        [HttpPost]
        public ActionResult Delete(int id, FormCollection collection)
        {
            try
            {
                // TODO: Add delete logic here

                return RedirectToAction("Index");
            }
            catch
            {
                return View();
            }
        }
    }
}
